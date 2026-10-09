"""
Flask-compatibility helpers shared by application.py and the Products/* routers.

Keeps every route body identical to the Flask version:
    jsonify, redirect, url_for, render_template, send_file
Place this file next to application.py (same folder as db_config.py).
"""
import json
import mimetypes
import os
import re
import sys
from datetime import date, datetime, timezone
from decimal import Decimal
from email.utils import format_datetime
from urllib.parse import quote, urlencode

from fastapi.responses import JSONResponse, RedirectResponse, Response
from starlette.routing import NoMatchFound
from fastapi.templating import Jinja2Templates

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, 'templates'))
_app = None


def init_app(app):
    """Call once from application.py right after `app = FastAPI()` (optional: the app is also auto-detected)."""
    global _app
    _app = app
    templates.env.globals['url_for'] = url_for


def _get_app():
    """The FastAPI app: set by init_app(), or found among the loaded modules."""
    global _app
    if _app is None:
        from fastapi import FastAPI
        for mod in list(sys.modules.values()):
            for value in list(getattr(mod, '__dict__', {}).values()):
                if isinstance(value, FastAPI):
                    _app = value
                    return _app
    return _app


def _json_default(o):
    """Same fallbacks Flask's jsonify used (Decimal -> str, date/datetime -> HTTP date)."""
    if isinstance(o, datetime):
        o = o.replace(tzinfo=timezone.utc) if o.tzinfo is None else o.astimezone(timezone.utc)
        return format_datetime(o, usegmt=True)
    if isinstance(o, date):
        return format_datetime(datetime(o.year, o.month, o.day, tzinfo=timezone.utc), usegmt=True)
    if isinstance(o, Decimal):
        return str(o)
    return str(o)


class FlaskJSONResponse(JSONResponse):
    def render(self, content) -> bytes:
        return json.dumps(content, default=_json_default, sort_keys=True,
                          ensure_ascii=True, separators=(',', ':')).encode('utf-8')


def jsonify(data=None, status_code=200):
    return FlaskJSONResponse(content=data, status_code=status_code)


def redirect(location, code=302):
    return RedirectResponse(url=location, status_code=code)


def _walk_routes(routes, prefix=''):
    """Yield (full_path, route) for every route, descending into nested/included routers."""
    for r in routes:
        children = getattr(r, 'routes', None)
        if children is None:
            children = getattr(getattr(r, 'router', None), 'routes', None)
        if children is None:
            children = getattr(getattr(r, 'original_router', None), 'routes', None)
        own = getattr(r, 'path', '') or ''
        if children is not None:
            yield from _walk_routes(children, prefix + (getattr(r, 'prefix', None) or own))
        else:
            yield prefix + own, r


_ROUTER_INDEX = None


def _router_index():
    """name -> path for every APIRouter object loaded in this process (prod_app, unit_app, fin_app, ...).
    Used when the app's own route tree does not expose included routes."""
    global _ROUTER_INDEX
    if _ROUTER_INDEX is None:
        from fastapi import APIRouter
        index = {}
        for mod in list(sys.modules.values()):
            for value in list(getattr(mod, '__dict__', {}).values()):
                if isinstance(value, APIRouter):
                    for r in value.routes:
                        nm = getattr(r, 'name', None)
                        if nm and nm not in index:
                            index[nm] = getattr(r, 'path', '')
        _ROUTER_INDEX = index
    return _ROUTER_INDEX


def url_for(endpoint, **values):
    """Flask style url_for: 'static' + filename, 'blueprint.endpoint' names, extra values -> query string."""
    if endpoint == 'static':
        filename = values.get('filename', values.get('path', ''))
        return '/static/' + str(filename).lstrip('/')
    name = endpoint.split('.')[-1]
    values.pop('_external', None)
    app = _get_app()
    if app is None:
        raise RuntimeError("No FastAPI app found; call fastapi_compat.init_app(app)")

    # 1) Walk the route tree (descends into routers included with app.include_router)
    for full_path, route in _walk_routes(app.routes):
        if getattr(route, 'name', None) == name \
                or getattr(getattr(route, 'endpoint', None), '__name__', None) == name:
            path_params = set(re.findall(r'{(\w+)', full_path))
            path_values = {k: values.pop(k) for k in list(values) if k in path_params}
            url = full_path.format(**path_values) if path_params else full_path
            return url + ('?' + urlencode(values, doseq=True) if values else '')

    # 2) Fallback: ask the framework itself (newer FastAPI/Starlette versions may keep included
    #    routers nested behind objects the walk above cannot read).
    for path_keys in ([], list(values)):
        try:
            url = str(app.url_path_for(name, **{k: values[k] for k in path_keys}))
        except NoMatchFound:
            continue
        rest = {k: v for k, v in values.items() if k not in path_keys}
        return url + ('?' + urlencode(rest, doseq=True) if rest else '')

    # 3) Last resort: routes defined on a router that the app tree does not expose
    path = _router_index().get(name)
    if path is not None:
        kinds = sorted({type(r).__name__ for r in app.routes})
        print(f"[fastapi_compat] '{name}' is defined on a router but is not visible in app.routes "
              f"(entry types: {kinds}); using the router's own path '{path}'")
        path_params = set(re.findall(r'{(\w+)', path))
        path_values = {k: values.pop(k) for k in list(values) if k in path_params}
        url = path.format(**path_values) if path_params else path
        return url + ('?' + urlencode(values, doseq=True) if values else '')
    raise KeyError(f"No route named '{endpoint}'")


templates.env.globals['url_for'] = url_for   # replaces Starlette's url_for (needs no init_app call)


def render_template(request, template_name, **context):
    global _app
    if _app is None:
        _app = request.app
    context.setdefault('session', request.session)
    return templates.TemplateResponse(request, template_name, context)


def send_file(path_or_file, mimetype=None, as_attachment=False, download_name=None):
    if isinstance(path_or_file, (str, os.PathLike)):
        with open(path_or_file, 'rb') as fh:
            data = fh.read()
        if download_name is None:
            download_name = os.path.basename(path_or_file)
    elif hasattr(path_or_file, 'getvalue'):
        data = path_or_file.getvalue()
    else:
        data = path_or_file.read()
    if mimetype is None and download_name:
        mimetype = mimetypes.guess_type(download_name)[0]
        if mimetype is None and download_name.lower().endswith('.xlsx'):
            mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    if mimetype is None:
        mimetype = 'application/octet-stream'
    headers = {'Content-Length': str(len(data))}
    if as_attachment and download_name:
        ascii_name = download_name.encode('ascii', 'ignore').decode().replace('"', '')
        headers['Content-Disposition'] = (
            f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{quote(download_name, safe='')}")
    return Response(content=data, media_type=mimetype, headers=headers)
