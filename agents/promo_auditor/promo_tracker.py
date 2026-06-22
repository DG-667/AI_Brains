import os
import yaml
from pathlib import Path

OBSIDIAN_VAULT_PATH = r"C:\Users\DenisG\Documents\Workplace\MyStorage\promo"

def create_promo_note(promo_name, brand, part_numbers, plan_qty=0):
    """Создает новую Markdown-заметку в Obsidian для акции с YAML Frontmatter."""
    if not os.path.exists(OBSIDIAN_VAULT_PATH):
        os.makedirs(OBSIDIAN_VAULT_PATH)
        
    safe_name = "".join(c for c in promo_name if c.isalnum() or c in " _-").strip()
    file_path = os.path.join(OBSIDIAN_VAULT_PATH, f"{safe_name}.md")
    
    # Свойства акции (метаданные для Dataview)
    frontmatter = {
        "status": "planned",
        "brand": brand,
        "part_numbers": part_numbers if isinstance(part_numbers, list) else [part_numbers],
        "plan_qty": plan_qty,
        "fact_qty": 0
    }
    
    # Формируем тело заметки
    content = f"---\n{yaml.dump(frontmatter, allow_unicode=True, default_flow_style=False)}---\n\n"
    content += f"# Акция: {promo_name}\n\n"
    content += "Эта заметка создана автоматически скриптом Promo Tracker.\n\n"
    content += "## Детали\n"
    content += f"- **Бренд:** {brand}\n"
    content += f"- **Модели (Part Numbers):** {', '.join(frontmatter['part_numbers'])}\n"
    content += f"- **План:** {plan_qty} шт.\n"
    content += f"- **Факт:** {frontmatter['fact_qty']} шт. *(будет обновляться агентом)*\n"
    
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(content)
        
    print(f"SUCCESS: Успешно создана заметка в Obsidian для акции '{promo_name}'. Файл: {file_path}")

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Obsidian Promo Tracker")
    parser.add_argument("action", choices=["test"], help="Действие")
    
    args = parser.parse_args()
    
    if args.action == "test":
        create_promo_note("Test Promo Lenovo", "Lenovo", ["82XQ015BRK", "82VG00TNRK"], 150)
