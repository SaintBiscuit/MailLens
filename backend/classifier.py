import torch
import numpy as np
import os
import random
from sentence_transformers import SentenceTransformer


class MailClassifier:
    def __init__(self, threshold=0.28):
        self.threshold = threshold
        if torch.cuda.is_available():
            torch.cuda.empty_cache()
            self.device = "cuda" 
        else:
            self.device = "cpu"
        self.model = SentenceTransformer(
            "intfloat/multilingual-e5-large",
            device=self.device
        )
        self.categories = {}

        self.category_prefix = "Категория писем:" 
        self.email_prefix = "Классифицируй это письмо:"  

    
    def add_category(self, category: str, description: str = '', example_texts: list[str] = None):
        """
        Добавляет категорию С ПРИМЕРАМИ писем
        
        Args:

        """
        if category:
            prompts = []
            if description:
                prompts.append(f"{self.category_prefix} {category}. Конкретное описание категории: {description}")
            else:
                prompts.append(f"{self.category_prefix} {category}")
            
            # 2. Добавляем примеры (если есть, максимум 5)
            if example_texts:
                for i, example in enumerate(example_texts[:5]):
                    prompt = f"Пример письма категории '{category}' #{i+1} (уникальные черты): {example}..."
                    prompts.append(prompt)
            
            # 3. Кодируем ВСЕ промпты категории в эмбеддинги
            category_embeddings = self.model.encode(
                prompts,
                normalize_embeddings=True,
                convert_to_numpy=True,
                show_progress_bar=False
            )
            # 4. Сохраняем категорию со всей информацией
            category_data = {
                'embeddings': category_embeddings,  # все эмбеддинги этой категории
            }
            self.categories[category] = category_data
        else:
            raise ValueError("Не была передана категория")

    def predict(self, text):
        if not self.categories:
            return {"error": "Нет категорий для классификации!"}
        
        mail_emb = self.model.encode(f"{self.email_prefix} {text}", normalize_embeddings=True)
        category_scores = {}
        
        # Для каждой категории считаем близость с её эмбеддингами
        for category, category_data in self.categories.items():
            cat_embeddings = category_data['embeddings']
            
            # Считаем косинусную близость со всеми эмбеддингами категории
            similarities = np.dot(cat_embeddings, mail_emb.T).flatten()
            
            # Выбираем МАКСИМАЛЬНУЮ близость
            category_scores[category] = float(np.max(similarities))
        
        # Нормализуем в вероятности (softmax)
        similarities = np.array(list(category_scores.values()))
        
        # Формируем результаты
        results = []
        for i, (category, similarity) in enumerate(category_scores.items()):
            results.append({
                "category": category,
                "similarity": float(similarity),
            })
        
        results.sort(key=lambda x: x["similarity"], reverse=True)
        
        # Применяем порог
        best_result = results[0]
        if best_result["similarity"] < self.threshold:
            predicted = "Не определена"
        else:
            predicted = best_result["category"]
        
        return {
            "predicted_category": predicted,
            "best_similarity": float(best_result["similarity"]),
            "all_scores": results,
        }
    
