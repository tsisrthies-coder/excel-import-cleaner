import datetime
import re

def parse_and_format_date(value: str) -> str:
    """
    Tente de parser différents formats de date et les renvoie TOUJOURS en JJ/MM/AAAA.
    Gère les séparateurs : /, -, .
    """
    val_str = value.strip()
    
    # 1. Détection format : JJ/MM/AAAA ou JJ-MM-AAAA ou JJ.MM.AAAA
    match_classique = re.match(r'^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$', val_str)
    if match_classique:
        d, m, y = match_classique.groups()
        return f"{int(d):02d}/{int(m):02d}/{int(y)}"
        
    # 2. Détection format ISO : AAAA-MM-JJ ou AAAA/MM/JJ
    match_iso = re.match(r'^(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})$', val_str)
    if match_iso:
        y, m, d = match_iso.groups()
        return f"{int(d):02d}/{int(m):02d}/{int(y)}"
        
    return val_str

def is_text_date(value: str) -> bool:
    """Vérifie si la chaîne de caractères ressemble à n'importe quel format de date"""
    if not isinstance(value, str):
        return False
    val_str = value.strip()
    
    pattern_classique = r'^\d{1,2}[/.-]\d{1,2}[/.-]\d{4}$'
    pattern_iso = r'^\d{4}[/.-]\d{1,2}[/.-]\d{1,2}$'
    
    return bool(re.match(pattern_classique, val_str)) or bool(re.match(pattern_iso, val_str))

def format_text_date(value: str) -> str:
    """Passe par la fonction de parsing universelle"""
    if isinstance(value, str):
        return parse_and_format_date(value)
    return value

def convert_to_string_date(value) -> str:
    """Convertit un objet datetime/date (vraie date Excel) en chaîne JJ/MM/AAAA"""
    if isinstance(value, (datetime.date, datetime.datetime)):
        return value.strftime('%d/%m/%Y')
    return str(value)