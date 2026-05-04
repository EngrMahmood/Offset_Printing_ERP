from pathlib import Path
text = Path('requirements.txt').read_text(encoding='utf-16')
print(text)
