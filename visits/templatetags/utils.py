from django import template
register = template.Library()

@register.filter
def get_item(d, key):
    try:
        return d.get(key)
    except:
        return None
