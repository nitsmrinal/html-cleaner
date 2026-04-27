import re
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from bs4 import BeautifulSoup, Comment
import uvicorn

app = FastAPI()


class HTMLInput(BaseModel):
    html: str


class HTMLOutput(BaseModel):
    html: str


# ---------------------------------------------------------------------------
# MSO → standard CSS conversion map
# Each key is a regex pattern matching the mso property name (case-insensitive).
# The value is either:
#   - a string: the standard CSS property to rename it to
#   - None: drop the property entirely (no CSS equivalent)
# ---------------------------------------------------------------------------
MSO_TO_CSS = {
    # Font / text
    r'mso-bidi-font-weight':        'font-weight',
    r'mso-bidi-font-style':         'font-style',
    r'mso-bidi-font-size':          'font-size',
    r'mso-bidi-font-family':        'font-family',
    r'mso-font-kerning':            'font-kerning',
    r'mso-font-width':              None,

    # Spacing / layout
    r'mso-margin-top-alt':          'margin-top',
    r'mso-margin-bottom-alt':       'margin-bottom',
    r'mso-margin-left-alt':         'margin-left',
    r'mso-margin-right-alt':        'margin-right',
    r'mso-padding-alt':             'padding',
    r'mso-line-height-alt':         'line-height',
    r'mso-line-height-rule':        None,
    r'mso-char-indent-count':       None,

    # Text decoration / color
    r'mso-color-alt':               'color',
    r'mso-highlight':               'background-color',
    r'mso-text-raise':              'vertical-align',

    # Table
    r'mso-table-lspace':            None,
    r'mso-table-rspace':            None,
    r'mso-cellspacing':             None,
    r'mso-border-alt':              'border',
    r'mso-border-top-alt':          'border-top',
    r'mso-border-bottom-alt':       'border-bottom',
    r'mso-border-left-alt':         'border-left',
    r'mso-border-right-alt':        'border-right',
}

# Compile patterns once at startup
_MSO_PATTERNS = [(re.compile(k, re.IGNORECASE), v) for k, v in MSO_TO_CSS.items()]

# MS-specific class prefixes to strip
_MS_CLASS_RE = re.compile(r'^(Mso|mso|xl\d)', re.IGNORECASE)

# Namespace tag pattern (e.g. o:p, w:view, v:shape)
_NS_TAG_RE = re.compile(r'^[a-z]+:', re.IGNORECASE)

# Hard-coded MS tag names to remove entirely
_MS_TAGS = {
    'o:p', 'o:smarttagtype', 'o:documentproperties', 'o:officedocumentsettings',
    'w:worddocument', 'w:view', 'w:zoom', 'w:donotoptimizeforbrowser',
    'x:excelworkbook', 'xml', 'style',
}


def _convert_mso_property(prop: str, value: str):
    """
    Convert a single mso-* CSS property+value to a standard CSS declaration.
    Returns "property:value" string, or None if it should be dropped.
    """
    prop = prop.strip()
    value = value.strip()

    for pattern, target in _MSO_PATTERNS:
        if pattern.fullmatch(prop):
            if target is None:
                return None  # explicitly dropped, no CSS equivalent
            return f"{target}:{value}"

    # Any remaining unrecognised mso-* property → drop
    if prop.lower().startswith('mso-'):
        return None

    # Non-mso property → keep unchanged
    return f"{prop}:{value}"


def _clean_style(style_str: str) -> str:
    """
    Process an inline style string:
    - Convert known mso-* properties to standard CSS equivalents
    - Drop all remaining unknown mso-* properties
    - Keep every other CSS property unchanged
    """
    declarations = [d.strip() for d in style_str.split(';') if d.strip()]
    result = []

    for decl in declarations:
        if ':' not in decl:
            continue
        prop, _, value = decl.partition(':')
        converted = _convert_mso_property(prop, value)
        if converted:
            result.append(converted)

    return '; '.join(result)


def clean_microsoft_html(html: str) -> str:
    """
    Convert Microsoft Office HTML to clean HTML with standard inline CSS.

    Steps:
    1. Strip XML declarations and xmlns namespace attributes
    2. Remove MS Office conditional comments
    3. Remove all MS-specific / namespaced tags (o:p, w:*, v:*, etc.)
    4. Convert mso-* inline styles → standard CSS equivalents
    5. Strip MS-specific class names (MsoNormal, etc.) and attributes
    6. Unwrap empty span/font wrappers left behind
    """

    # 1. Strip XML declaration and xmlns attributes
    html = re.sub(r'<\?xml[^>]*\?>', '', html, flags=re.IGNORECASE)
    html = re.sub(r'\s+xmlns[:a-z]*="[^"]*"', '', html, flags=re.IGNORECASE)

    # 2. Remove MS conditional comments (<!--[if ...]> ... <![endif]-->)
    html = re.sub(r'<!--\[if[^\]]*\]>.*?<!\[endif\]-->', '', html, flags=re.DOTALL | re.IGNORECASE)
    html = re.sub(r'<!--\[if[^\]]*\]>.*?-->', '', html, flags=re.DOTALL | re.IGNORECASE)

    soup = BeautifulSoup(html, 'html.parser')

    # Remove plain HTML comments
    for comment in soup.find_all(string=lambda text: isinstance(text, Comment)):
        comment.extract()

    # 3. Remove MS-specific and namespaced tags (and all their children)
    for tag in soup.find_all(True):
        if tag.name in _MS_TAGS or _NS_TAG_RE.match(tag.name):
            tag.decompose()

    # 4 & 5. Process attributes on every remaining tag
    for tag in soup.find_all(True):
        attrs_to_remove = []

        for attr in list(tag.attrs.keys()):

            # style → convert mso-* to standard CSS, keep everything else
            if attr == 'style':
                cleaned = _clean_style(tag['style'])
                if cleaned:
                    tag['style'] = cleaned
                else:
                    attrs_to_remove.append('style')

            # class → strip MS class names, keep custom ones
            elif attr == 'class':
                classes = tag['class']
                if isinstance(classes, str):
                    classes = classes.split()
                kept = [c for c in classes if not _MS_CLASS_RE.match(c)]
                if kept:
                    tag['class'] = kept
                else:
                    attrs_to_remove.append('class')

            # lang / xml:lang → always remove
            elif attr in ('lang', 'xml:lang'):
                attrs_to_remove.append(attr)

            # any remaining namespaced attribute (e.g. v:shapes) → remove
            elif ':' in attr:
                attrs_to_remove.append(attr)

        for attr in attrs_to_remove:
            del tag[attr]

    # 6. Remove or unwrap empty span/font wrappers with no remaining attributes
    for tag in soup.find_all(['span', 'font']):
        if not tag.attrs:
            if not tag.contents:
                tag.decompose()
            else:
                tag.unwrap()

    result = str(soup)
    result = re.sub(r'\n{3,}', '\n\n', result)
    return result.strip()


@app.post("/clean", response_model=HTMLOutput)
async def clean_html(input_data: HTMLInput):
    """
    Convert Microsoft Office HTML to clean, standard HTML.

    What this does:
    - Replaces mso-* CSS properties with standard inline CSS equivalents
    - Strips all MS-specific tags: o:p, w:*, x:*, v:*, xml, style blocks
    - Removes MS class names (MsoNormal, MsoBodyText, etc.)
    - Drops lang, xml:lang, and other MS-injected attributes
    - Preserves all legitimate formatting (color, font-size, font-family, etc.)
    """
    if not input_data.html:
        raise HTTPException(status_code=400, detail="HTML input cannot be empty")

    cleaned = clean_microsoft_html(input_data.html)
    return HTMLOutput(html=cleaned)


if __name__ == "__main__":
    uvicorn.run("file:app", host="0.0.0.0", port=8000, reload=True)
