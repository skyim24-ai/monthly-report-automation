import re
import os

class DesignConfig:
    def __init__(self, file_path="design.md"):
        self.settings = {
            "pptx": {
                "title_size": 44,
                "body_size": 12,
                "title_color": "000000",
                "accent_color": "0078D4",
                "font_family": "맑은 고딕"
            },
            "docx": {
                "title_size": 36,
                "heading1_size": 18,
                "body_size": 11,
                "font_family": "맑은 고딕",
                "text_color": "000000"
            }
        }
        if os.path.exists(file_path):
            self.parse_md(file_path)

    def parse_md(self, path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                content = f.read()

            sections = re.split(r'##\s*', content)
            for section in sections:
                lines = section.strip().split('\n')
                if not lines: continue
                
                section_name = lines[0].lower()
                target = None
                if "pptx" in section_name:
                    target = self.settings["pptx"]
                elif "docx" in section_name:
                    target = self.settings["docx"]
                
                if target:
                    for line in lines[1:]:
                        match = re.search(r'-\s*(.*?):\s*(.*)', line)
                        if match:
                            key, val = match.groups()
                            key = self.map_key(key.strip())
                            if key:
                                target[key] = self.format_val(val.strip())
        except Exception as e:
            print(f"Design parsing error: {e}")

    def map_key(self, key):
        mapping = {
            "제목 폰트 크기": "title_size",
            "제목 크기": "title_size",
            "본문 폰트 크기": "body_size",
            "본문 크기": "body_size",
            "제목 색상": "title_color",
            "글자 색상": "text_color",
            "글자 색": "text_color",
            "강조 색상": "accent_color",
            "폰트 종류": "font_family",
            "폰트": "font_family",
            "대제목 크기": "title_size",
            "소제목 크기": "heading1_size"
        }
        return mapping.get(key)

    def format_val(self, val):
        # Hex color check
        if val.startswith('#'):
            return val.replace('#', '')
        # Number check
        if val.isdigit():
            return int(val)
        return val

    def get(self, section, key):
        return self.settings.get(section, {}).get(key)
