def detect_injection(text: str):
    dangerous = ["ignore", "забудь", "ты теперь", "system prompt", "инъекция"]
    for word in dangerous:
        text.replace(word, '')
    return text