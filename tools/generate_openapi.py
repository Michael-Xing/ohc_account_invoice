#!/usr/bin/env python3
"""
Generate OpenAPI JSON and a static Swagger UI HTML file.

Usage:
    python tools/generate_openapi.py

This imports the FastAPI app from `src.main:app`, writes `generated_files/openapi/openapi.json`
and `generated_files/openapi/swagger.html`.
"""
from pathlib import Path
import json
import sys
import os

# Ensure project root on path
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

try:
    # Skip infrastructure initialization for OpenAPI generation
    os.environ.setdefault("SKIP_INFRA_INIT", "1")
    # 配置现在从 config/config.toml 文件加载，不再需要设置环境变量

    from src.main import app  # import the FastAPI app
except Exception as exc:
    print("Failed to import FastAPI app from src.main:", exc)
    raise


def main() -> int:
    out_dir = ROOT / "src" / "swagger"
    out_dir.mkdir(parents=True, exist_ok=True)

    # Export OpenAPI JSON
    openapi = app.openapi()
    openapi_path = out_dir / "openapi.json"
    with open(openapi_path, "w", encoding="utf-8") as f:
        json.dump(openapi, f, ensure_ascii=False, indent=2)
    print("Wrote OpenAPI JSON to", openapi_path)

    # Write a simple Swagger UI HTML that loads the JSON from same directory
    swagger_html = out_dir / "swagger.html"
    html_content = f"""<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8"/>
    <title>Swagger UI</title>
    <link rel="stylesheet" href="https://unpkg.com/swagger-ui-dist@4/swagger-ui.css" />
  </head>
  <body>
    <div id="swagger-ui"></div>
    <script src="https://unpkg.com/swagger-ui-dist@4/swagger-ui-bundle.js"></script>
    <script>
      window.onload = function() {{
        const ui = SwaggerUIBundle({{
          url: 'openapi.json',
          dom_id: '#swagger-ui',
        }});
      }};
    </script>
  </body>
</html>"""

    with open(swagger_html, "w", encoding="utf-8") as f:
        f.write(html_content)
    print("Wrote static Swagger UI to", swagger_html)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


