def convert_to_mm(dim_str):
    """Converte dimensões em polegadas (ex: "1.1/2"") para mm."""
    dim_str = dim_str.strip().replace(',', '.')
    total_mm = 0.0
    try:
        if '"' in dim_str:
            dim_str = dim_str.replace('"', '')
            parts = dim_str.split('.')
            if parts[0] and '/' in parts[0]:
                num, den = map(float, parts[0].split('/'))
                total_mm += (num / den) * 25.4
            elif parts[0]:
                total_mm += float(parts[0]) * 25.4
            if len(parts) > 1 and '/' in parts[1]:
                num, den = map(float, parts[1].split('/'))
                total_mm += (num / den) * 25.4
        else:
            total_mm = float(dim_str)
    except (ValueError, ZeroDivisionError): return 0.0
    return total_mm


def normalizar_viga_w(nome_bruto):
    """
    Converte 'W 200 46.1' ou 'W 200 x 46.1' em 'W200X46,1'
    para bater com o padrão da coluna A do seu Excel.
    """
    if not nome_bruto: return ""
    # Remove espaços, coloca em maiúsculo e troca ponto por vírgula
    limpo = nome_bruto.upper().replace(" ", "").replace("X", "")
    # Reconstrói com o 'X' e a vírgula decimal brasileira
    # Exemplo: de W20046.1 para W200X46,1
    if "W" in limpo:
        # Pega os números após o W
        partes = limpo.replace("W", "")
        # Esta lógica assume que o primeiro grupo de números é a altura (ex: 200)
        # Se os nomes no Word variarem muito, podemos usar Regex aqui.
        return f"W{partes}".replace(".", ",") 
    return limpo

def cm_para_m(valor_cm):
    """Converte os valores da lista (kgf-cm) para metros."""
    try:
        return float(valor_cm) / 100
    except (ValueError, TypeError):
        return 0.0
    
def normalizar_nome_para_comparacao(texto):
    """Remove espaços, troca ponto por vírgula e garante o 'X'."""
    if not texto: return ""
    # Ex: 'W 200 46.1' -> 'W20046,1'
    t = str(texto).upper().replace(" ", "").replace(".", ",")
    # Se for viga W e não tiver o X, tentamos padronizar (ex: W20046,1 -> W200X46,1)
    if t.startswith('W') and 'X' not in t:
        # Lógica simples: insere o X após o primeiro grupo de 3 números
        # Ou simplesmente compare ignorando o 'X'
        return t.replace("W", "W") 
    return t