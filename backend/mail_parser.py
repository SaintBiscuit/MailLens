import email
from email import policy
import re
import html
import extract_msg
from typing import Dict, Any
import charset_normalizer


def guess_mime_type(filename: str) -> str:
    """Определяет MIME тип по расширению файла"""
    ext = filename.lower().split('.')[-1] if '.' in filename else ''
    
    mime_map = {
        'pdf': 'application/pdf',
        'doc': 'application/msword',
        'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'xls': 'application/vnd.ms-excel',
        'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'ppt': 'application/vnd.ms-powerpoint',
        'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'txt': 'text/plain',
        'zip': 'application/zip',
        'rar': 'application/x-rar-compressed',
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'png': 'image/png',
        'gif': 'image/gif',
    }

    return mime_map.get(ext, 'application/octet-stream')

def parse_email(file_bytes: bytes, filename: str = "unknown") -> Dict[str, Any]:
    """
    РАБОЧИЙ парсер писем с поддержкой HTML и правильной кодировкой
    """
    subject = "Без темы"
    text_parts = []
    html_parts = []
    attachments = []
    links = []
    
    def decode_payload(payload_bytes):
        """Определяет кодировку и декодирует"""
        if not payload_bytes:
            return ""
        
        # Используем charset-normalizer для определения кодировки
        result = charset_normalizer.from_bytes(payload_bytes).best()
        if result:
            return str(result)
        
        # Пробуем стандартные кодировки
        for encoding in ['utf-8', 'cp1251', 'koi8-r', 'iso-8859-5', 'windows-1251']:
            try:
                return payload_bytes.decode(encoding, errors='ignore')
            except:
                continue
        
        # Последняя попытка
        try:
            return payload_bytes.decode('utf-8', errors='ignore')
        except:
            return ""
    
    def html_to_text(html_content):
        """Конвертация HTML в текст"""
        if not html_content:
            return ""
        
        # Убираем теги
        text = re.sub(r'<[^>]+>', ' ', html_content)
        
        # Заменяем HTML entities
        text = html.unescape(text)
        
        # Убираем лишние пробелы
        text = re.sub(r'\s+', ' ', text)
        
        return text.strip()
    
    try:
        # === .MSG ФАЙЛЫ ===
        if filename.lower().endswith('.msg'):
            try:
                msg = extract_msg.Message(file_bytes)
                subject = str(msg.subject) if msg.subject else "Без темы"
                
                # Получаем тело
                body = msg.body if hasattr(msg, 'body') else ""
                if body:
                    # Проверяем, HTML ли это
                    if isinstance(body, bytes):
                        body = decode_payload(body)
                    
                    if '<html' in body.lower() or '<body' in body.lower():
                        html_parts.append(body)
                        # Конвертируем HTML в текст
                        plain_text = html_to_text(body)
                        if plain_text:
                            text_parts.append(plain_text)
                    else:
                        text_parts.append(body)
                
                # Вложения
                if hasattr(msg, 'attachments'):
                    for att in msg.attachments:
                        if hasattr(att, 'longFilename') and att.longFilename:
                            attachments.append({
                                'filename': str(att.longFilename),
                                'type': guess_mime_type(str(att.longFilename))
                            })
            except Exception as e:
                print(f"Ошибка парсинга .msg: {e}")
                text_parts.append("[Ошибка чтения .msg файла]")
        
        # === .EML ФАЙЛЫ ===
        else:
            try:
                msg = email.message_from_bytes(file_bytes, policy=policy.default)
                subject = str(msg.get('Subject', 'Без темы'))
                
                def process_part(part):
                    """Рекурсивно обрабатывает части письма"""
                    content_type = part.get_content_type().lower()
                    content_disposition = str(part.get('Content-Disposition', '')).lower()
                    
                    # Если это вложение
                    if 'attachment' in content_disposition or part.get_filename():
                        filename = part.get_filename()
                        if filename:
                            attachments.append({
                                'filename': filename,
                                'type': content_type
                            })
                        return
                    
                    # Получаем содержимое
                    payload = part.get_payload(decode=True)
                    if not payload:
                        return
                    
                    decoded = decode_payload(payload)
                    if not decoded:
                        return
                    
                    # Обрабатываем по типу
                    if content_type == 'text/plain':
                        if decoded.strip():
                            text_parts.append(decoded.strip())
                    
                    elif content_type == 'text/html':
                        html_parts.append(decoded)
                        # Конвертируем HTML в текст
                        plain_text = html_to_text(decoded)
                        if plain_text.strip():
                            text_parts.append(plain_text.strip())
                    
                    # Рекурсивно обрабатываем multipart
                    elif content_type.startswith('multipart/'):
                        for subpart in part.get_payload():
                            if isinstance(subpart, email.message.Message):
                                process_part(subpart)
                
                # Запускаем обработку
                if msg.is_multipart():
                    for part in msg.walk():
                        process_part(part)
                else:
                    process_part(msg)
                    
            except Exception as e:
                print(f"Ошибка парсинга .eml: {e}")
                text_parts.append("[Ошибка чтения .eml файла]")
        
        all_text = []
        
        for text in text_parts:
            if text and text.strip():
                all_text.append(text.strip())
        
        if not all_text and html_parts:
            for html_content in html_parts:
                text = html_to_text(html_content)
                if text.strip():
                    all_text.append(text.strip())
        
        if not all_text:
            all_text = ["[Пустое письмо]"]
        
        full_text = "\n\n".join(all_text)
        
        url_pattern = r'https?://[^\s<>"]+|www\.[^\s<>"]+'
        urls = re.findall(url_pattern, full_text, re.IGNORECASE)
        
        for html_content in html_parts:
            html_urls = re.findall(r'href=[\'"]?([^\'" >]+)', html_content, re.IGNORECASE)
            urls.extend(html_urls)
        
        unique_urls = set()
        for url in urls:
            if url and url not in unique_urls:
                unique_urls.add(url)
                clean_url = url.split('#')[0].split('?')[0]
                if clean_url.startswith(('http://', 'https://', 'www.')):
                    match = re.search(r'https?://(?:www\.)?([^/\?:]+)', clean_url, re.IGNORECASE)
                    if match:
                        links.append({
                            'url': clean_url,
                            'domain': match.group(1).lower()
                        })
        
        clean_body = full_text
        
        clean_body = re.sub(r'\n{3,}', '\n\n', clean_body)
        clean_body = re.sub(r'[ \t]{2,}', ' ', clean_body)
        return {
            "subject": subject,
            "body": clean_body.strip(),
            "has_attachments": len(attachments) > 0,
            "attachments": attachments,
            "links": links,
            "full_raw_text": full_text,
            "file_type": filename.split('.')[-1].lower() if '.' in filename else "txt",
            "text_parts_count": len(text_parts),
            "html_parts_count": len(html_parts)
        }
    
    except Exception as e:
        raise Exception(f'Ошибка парсинга письма: {filename}')

def prepare_for_classification(parsed_data: Dict[str, Any], max_length: int = 1500) -> str:
    """
    Готовит текст для модели.
    """
    parts = []

    subject = parsed_data.get("subject", "")
    if subject and subject != "Без темы":
        parts.append(f"Тема: {subject[:200]}")

    body = parsed_data.get("body", "")
    if body:
        body = re.sub(r'\s+', ' ', body)  
        body = body.strip()
        
        if body:
            parts.append(f"Текст: {body}")

    if parsed_data.get("has_attachments", False):
        attachments = parsed_data.get("attachments", [])
        file_types = set()
        for att in attachments[:3]:
            if isinstance(att, dict) and "type" in att:
                file_type = att["type"].split('/')[-1]
                file_types.add(file_type)
        
        if file_types:
            parts.append(f"Вложения: {', '.join(sorted(file_types))}")
        else:
            parts.append("Есть вложения")

    links = parsed_data.get("links", [])
    if links:
        domains = set()
        for link in links[:5]:
            if isinstance(link, dict) and "domain" in link:
                domains.add(link["domain"])
        
        if domains:
            parts.append(f"Ссылки на: {', '.join(sorted(domains)[:3])}")

    final_text = ". ".join(parts)

    if len(final_text) > max_length:
        final_text = final_text[:max_length] + "..."
    
    return final_text
