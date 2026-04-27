import re
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from bs4 import BeautifulSoup, Comment

app = FastAPI()


class HTMLInput(BaseModel):
    html: str


class HTMLOutput(BaseModel):
    html: str


def clean_microsoft_html(html: str) -> str:
    """Remove Microsoft Office-specific formatting from HTML."""
    
    # Remove XML namespaces and declarations
    html = re.sub(r'<\?xml[^>]*\?>', '', html, flags=re.IGNORECASE)
    html = re.sub(r'xmlns[:a-z]*="[^"]*"', '', html, flags=re.IGNORECASE)
    
    # Remove MS Office conditional comments
    html = re.sub(r'<!--\[if[^\]]*\]>.*?<!\[endif\]-->', '', html, flags=re.DOTALL | re.IGNORECASE)
    html = re.sub(r'<!--\[if[^\]]*\]>.*?-->', '', html, flags=re.DOTALL | re.IGNORECASE)
    
    soup = BeautifulSoup(html, 'html.parser')
    
    # Remove HTML comments
    for comment in soup.find_all(string=lambda text: isinstance(text, Comment)):
        comment.extract()
    
    # Remove MS-specific tags
    ms_tags = [
        'o:p', 'o:smarttagtype', 'o:documentproperties', 'o:officedocumentsettings',
        'w:worddocument', 'w:view', 'w:zoom', 'w:donotoptimizeforbrowser',
        'x:excellworkbook', 'xml', 'style'  # style often contains mso-* rules
    ]
    for tag_name in ms_tags:
        for tag in soup.find_all(tag_name):
            tag.decompose()
    
    # Also remove any tag with namespace prefix (e.g., v:shape, o:lock)
    for tag in soup.find_all(re.compile(r'^[a-z]+:')):
        tag.decompose()
    
    # Clean attributes from all elements
    ms_class_patterns = re.compile(r'^(Mso|mso|xl\d)', re.IGNORECASE)
    
    for tag in soup.find_all(True):
        # Remove MS-specific attributes
        attrs_to_remove = []
        for attr in list(tag.attrs.keys()):
            # Remove style attributes with mso-* properties
            if attr == 'style':
                style = tag.get('style', '')
                # Remove mso-* properties from inline styles
                cleaned_style = re.sub(r'mso-[^;:]+:[^;]+;?', '', style, flags=re.IGNORECASE)
                cleaned_style = re.sub(r';\s*;', ';', cleaned_style)  # Clean up double semicolons
                cleaned_style = cleaned_style.strip('; ')
                if cleaned_style:
                    tag['style'] = cleaned_style
                else:
                    attrs_to_remove.append('style')
            # Remove MS-specific class names
            elif attr == 'class':
                classes = tag.get('class', [])
                if isinstance(classes, str):
                    classes = classes.split()
                cleaned_classes = [c for c in classes if not ms_class_patterns.match(c)]
                if cleaned_classes:
                    tag['class'] = cleaned_classes
                else:
                    attrs_to_remove.append('class')
            # Remove lang attribute (often MS-specific like "EN-US")
            elif attr in ['lang', 'xml:lang']:
                attrs_to_remove.append(attr)
            # Remove any o:*, w:*, x:* attributes
            elif ':' in attr:
                attrs_to_remove.append(attr)
        
        for attr in attrs_to_remove:
            del tag[attr]
    
    # Remove empty spans and fonts that were just style wrappers
    for tag in soup.find_all(['span', 'font']):
        if not tag.attrs and tag.string is None:
            tag.unwrap()
    
    # Clean up excessive whitespace
    result = str(soup)
    result = re.sub(r'\n\s*\n', '\n\n', result)
    result = result.strip()
    
    return result


@app.post("/clean", response_model=HTMLOutput)
async def clean_html(input_data: HTMLInput):
    """
    Remove Microsoft-specific formatting from HTML.
    
    Strips:
    - MS Office conditional comments
    - o:p, w:*, x:* and other namespace tags
    - mso-* CSS properties
    - MsoNormal and similar class names
    - XML declarations and namespaces
    """
    if not input_data.html:
        raise HTTPException(status_code=400, detail="HTML input cannot be empty")
    
    cleaned = clean_microsoft_html(input_data.html)
    return HTMLOutput(html=cleaned)
