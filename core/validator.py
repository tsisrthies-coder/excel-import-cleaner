import os

def validate_file(filepath: str) -> tuple[bool, str]:
    if not filepath or not os.path.exists(filepath):
        return False, "Le fichier n'existe pas."
    
    ext = os.path.splitext(filepath)[1].lower()
    if ext not in ['.xlsx', '.xlsm']:
        return False, "Seuls les fichiers .xlsx et .xlsm sont supportés."
    
    try:
        # Vérifie si le fichier est ouvert/verrouillé par une autre application (ex: Excel)
        with open(filepath, 'a'):
            pass
    except PermissionError:
        return False, "Le fichier est ouvert. Veuillez le fermer dans Excel avant de continuer."
    except Exception as e:
        return False, f"Erreur d'accès : {str(e)}"
        
    return True, "Fichier valide."