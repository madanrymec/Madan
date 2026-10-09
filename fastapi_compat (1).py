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
from datetime import date, datetime, timezone
from decimal import Decimal
from email.utils import format_datetime
from urllib.parse import quote, urlencode

from fastapi.responses import JSONResponse, RedirectResponse, Response
from fastapi.templating import Jinja2Templates

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, 'templates'))
_app = None


def init_app(app):
    """Call once from application.py right after `app = FastAPI()`."""
    global _app
    _app = app
    templates.env.globals['url_for'] = url_for


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


def url_for(endpoint, **values):
    """Flask style url_for: 'static' + filename, 'blueprint.endpoint' names, extra values -> query string."""
    if endpoint == 'static':
        filename = values.get('filename', values.get('path', ''))
        return '/static/' + str(filename).lstrip('/')
    name = endpoint.split('.')[-1]
    values.pop('_external', None)
    if _app is None:
        raise RuntimeError("fastapi_compat.init_app(app) has not been called")
    matches = [r for r in _app.routes
               if getattr(r, 'name', None) == name
               or getattr(getattr(r, 'endpoint', None), '__name__', None) == name]
    if not matches:
        known = sorted({getattr(r, 'name', '') for r in _app.routes if getattr(r, 'name', None)})
        raise KeyError(f"No route named '{endpoint}' (registered: {len(known)} routes; "
                       f"is the router that defines '{name}' included with app.include_router?)")
    route = matches[0]
    path = getattr(route, 'path', None) or getattr(route, 'path_format', '')
    path_params = set(re.findall(r'{(\w+)', path))
    path_values = {k: values.pop(k) for k in list(values) if k in path_params}
    url = path.format(**path_values) if path_params else path
    return url + ('?' + urlencode(values, doseq=True) if values else '')


def render_template(request, template_name, **context):
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
