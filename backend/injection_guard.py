import re


def detect_injection(text: str, dangerous_patterns: list[str] = None) -> str:
    """
    Защита от prompt-инъекций с учётом разных вариаций
    
    Args:
        text: входной текст
        dangerous_patterns: список опасных паттернов
        
    Returns:
        Очищенный текст
    """
    if dangerous_patterns is None:
        dangerous_patterns = [
            # Английские паттерны
            r'ignore.*previous',
            r'disregard.*instructions',
            r'you are now.*assistant',
            r'new.*instructions',
            r'system.*prompt',
            r'ignore.*all',
            r'forget.*everything',
            
            # Русские паттерны
            r'забудь.*всё',
            r'забудь.*инструкции', 
            r'ты теперь.*помощник',
            r'с этого момента',
            r'игнорируй.*предыдущие',
            r'новые.*инструкции',
            r'выведи.*системный',
            
            # Общие
            r'prompt.*injection',
            r'инъекция.*промпта',
            r'взлом.*промпта',
            r'ignore.*above',
            r'output.*only'
        ]
    
    result = text
    
    for pattern in dangerous_patterns:
        # Игнорируем регистр, удаляем все вхождения
        result = re.sub(pattern, '', result, flags=re.IGNORECASE)
    
    return result.strip()