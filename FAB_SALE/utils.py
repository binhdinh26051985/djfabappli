from unittest import result
import io as StringIO
from io import BytesIO
from django.http import HttpResponse
from django.template.loader import get_template
from django.template import Context
from html import escape

from xhtml2pdf import pisa
'''
def html2pdf(template_src, context_dict={}):
    template = get_template(template_src)
    html  = template.render(context_dict)
    result = BytesIO()
    pdf = pisa.pisaDocument(BytesIO(html.encode("ISO-8859-1")), result)
    if not pdf.err:
        return HttpResponse(result.getvalue(), content_type='application/pdf')
    return None

'''

def render_to_pdf(template_src, context_dict={}):
    template = get_template(template_src)
    #context = Context(context_dict)
    html  = template.render(context_dict)
    #result = StringIO.StringIO()
    result = BytesIO()

    #pdf = pisa.pisaDocument(StringIO.StringIO(html.encode("ISO-8859-1")), result)
    pdf = pisa.pisaDocument(BytesIO(html.encode("ISO-8859-1")), result)
    if not pdf.err:
        return HttpResponse(result.getvalue(), content_type='application/pdf')
    #return HttpResponse('We had some errors<pre>%s</pre>' % escape(html))
    return None