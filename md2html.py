import markdown
import sys

md_file = sys.argv[1]
html_file = sys.argv[2]

with open(md_file, "r", encoding="utf-8") as f:
    text = f.read()

html = markdown.markdown(text, extensions=['tables'])

full_html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<style>
body {{ font-family: 'Microsoft JhengHei', sans-serif; padding: 40px; line-height: 1.6; }}
table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; }}
th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
th {{ background-color: #f2f2f2; }}
h1, h2, h3 {{ color: #333; }}
</style>
</head>
<body>
{html}
</body>
</html>"""

with open(html_file, "w", encoding="utf-8") as f:
    f.write(full_html)
