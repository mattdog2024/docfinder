import yaml

with open('.github/workflows/build-release.yml', 'r', encoding='utf-8') as f:
    content = f.read()

# 在 Install dependencies 后面插入生成图片的步骤
insert_str = """
      - name: Generate title image with Windows font
        run: |
          python -c "from PIL import Image, ImageDraw, ImageFont; import os; font = ImageFont.truetype('C:/Windows/Fonts/msyh.ttc', 22); img = Image.new('RGBA', (220, 40), (255,255,255,0)); draw = ImageDraw.Draw(img); draw.text((10, 8), '文档搜索索引', fill=(21,101,192,255), font=font); os.makedirs('assets', exist_ok=True); img.save('assets/title_text.png'); print('title_text.png generated')"
"""

if "Generate title image with Windows font" not in content:
    parts = content.split("      - name: Build EXE with PyInstaller")
    new_content = parts[0] + insert_str + "      - name: Build EXE with PyInstaller" + parts[1]
    
    with open('.github/workflows/build-release.yml', 'w', encoding='utf-8') as f:
        f.write(new_content)
    print("Workflow updated successfully.")
else:
    print("Workflow already updated.")
