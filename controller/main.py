from odoo import http
from odoo.http import request
from odoo.addons.web.controllers.main import serialize_exception,content_disposition
import base64


class Binary(http.Controller):

    @http.route('/web/binary/download_document/SO/<model("sale.exl"):order_line>', type='http', auth="user")
    @serialize_exception
    def download_document(self, order_line, **kw):
        res = order_line.print_excel_file()
        data = res and res[0] or ''
        filename = len(res)==2 and res[1] or 'Sale Order Lines'
        filecontent = base64.b64decode(data)
        if not filecontent:
            return request.not_found()
        else:
            return request.make_response(filecontent,
                            [('Content-Type', 'application/octet-stream'),
                             ('Content-Disposition', content_disposition(filename + '.xls'))])