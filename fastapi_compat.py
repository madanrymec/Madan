"""
Flask-compatibility helpers shared by application.py and the Products/* routers.

Keeps every route body identical to the Flask version:
    jsonify, redirect, url_for, render_template, send_file
Place this file next to application.py (same folder as db_config.py).
"""
import json
import mimetypes
import os
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
    """Flask style url_for: 'static' + filename, blueprint.endpoint names, extra values -> query string."""
    if endpoint == 'static':
        filename = values.get('filename', values.get('path', ''))
        return '/static/' + str(filename).lstrip('/')
    name = endpoint.split('.')[-1]
    values.pop('_external', None)
    for route in _app.routes:
        if getattr(route, 'name', None) == name and hasattr(route, 'param_convertors'):
            path_values = {k: values.pop(k) for k in list(values) if k in route.param_convertors}
            url = str(route.url_path_for(name, **path_values))
            return url + ('?' + urlencode(values, doseq=True) if values else '')
    raise KeyError(f"No route named '{endpoint}'")


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
