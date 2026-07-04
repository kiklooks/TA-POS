from django import template

register = template.Library()


@register.filter(name='rupiah')
def rupiah(value):
    try:
        number = int(float(value))
    except (TypeError, ValueError):
        return value

    return f"Rp{number:,.0f}".replace(",", ".")
