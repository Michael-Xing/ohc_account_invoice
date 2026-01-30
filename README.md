# OHC账票生成FastAPI服务

一个基于FastAPI框架的现代RESTful API服务，用于生成OHC账票文档。支持15种不同的Excel和Word模板填充，存储采用 MinIO。

## 功能特性

- 🚀 支持15种预定义账票模板
- 📊 Excel和Word文档模板填充
- 💾 存储选项：MinIO（默认）
- 🔧 基于uv的现代Python包管理
- 🐳 Docker容器化部署
- ⚡ FastAPI高性能框架
- 🔌 标准RESTful API支持
- 📚 自动生成的Swagger API文档
- 🎯 策略模式模板填充
- 🧠 智能内容识别和填充

## 支持的账票模板

1. **制作文档・图纸一览 (DHF INDEX)** - Excel
2. **PTF INDEX** - Excel
3. **ES个别试验要项书** - Excel
4. **ES个别试验结果书** - Excel/Word
5. **PP个别试验结果书** - Excel/Word
6. **ES验证计划书** - Word
7. **ES验证结果书** - Excel/Word
8. **PP验证计划书** - Word
9. **PP验证结果书** - Excel/Word
10. **基本规格书** - Word
11. **PP个别试验要项书** - Excel
12. **跟进DR会议记录** - Word
13. **标签规格书** - Word
14. **产品环境评估要项书/结果书** - Excel/Word
15. **与现有产品对比表** - Excel

## 快速开始

### 环境要求

- Python 3.11+
- uv (推荐) 或 pip
- Docker (可选)

### 安装依赖

```bash
# 使用uv (推荐) - 自动创建虚拟环境
make install

# 或手动使用uv
uv venv
uv pip install -e .
uv pip install -e ".[dev]"

# 或使用pip
pip install -e .
```

### 配置文件

编辑 `config/config.toml` 文件，示例如下（支持三种存储类型：`minio`、`local`、`s3`）：

```toml
[storage]
type = "minio"  # 可选值: minio, local, s3

[storage.minio]
endpoint = "localhost:9000"
access_key = "minioadmin"
secret_key = "minioadmin"
bucket_name = "ohc-documents"
secure = false

[storage.local]
path = "generated_files"

[storage.s3]
access_key_id = "your_aws_key"
secret_access_key = "your_aws_secret"
bucket_name = "ohc-documents"
region = "us-east-1"

[app]
host = "0.0.0.0"
port = 8000
debug = false
```

配置文件查找优先级：
1. `config/config.toml`（当前工作目录）
2. `config.toml`（当前工作目录）
3. `config/config.toml`（项目根目录）
4. `config.toml`（项目根目录）

也可以通过环境变量 `CONFIG_FILE` 指定自定义配置文件路径。

### 本地开发运行

```bash
# 使用Makefile - 开发模式（自动重载）
make dev

# 生产模式
make prod

# 直接运行
make run

# 查看API文档地址
make docs

# 或直接使用uv运行
uv run uvicorn src.main:app --host 0.0.0.0 --port 8000 --reload
```

### 访问API文档

启动服务后，可以通过以下地址访问API文档：

- **Swagger UI**: http://localhost:8000/docs
- **ReDoc**: http://localhost:8000/redoc
- **OpenAPI JSON**: http://localhost:8000/openapi.json

### API使用示例

```bash
# 1. 健康检查
curl http://localhost:8000/health

# 2. 获取模板列表
curl http://localhost:8000/templates

# 3. 获取模板信息
curl http://localhost:8000/templates/DHF_INDEX

# 4. 生成文档（通用接口，使用默认值）
curl -X POST "http://localhost:8000/generate" \
  -H "Content-Type: application/json" \
  -d '{
    "template_name": "DHF_INDEX",
    "parameters": {
      "project_name": "OHC项目",
      "version": "1.0",
      "document_type": "设计文档",
      "department": "研发部"
    }
  }'

# 5. 生成文档（专门接口）
curl -X POST "http://localhost:8000/generate/dhf-index" \
  -H "Content-Type: application/json" \
  -d '{
    "project_name": "OHC项目",
    "version": "1.0",
    "date": "2025-01-22",
    "author": "张三",
    "document_type": "设计文档",
    "department": "研发部",
    "reviewer": "李四",
    "approval_date": "2025-01-23"
  }'

# 6. 下载文件
使用生成接口返回的 presigned URL 或 MinIO 提供的下载链接进行文件下载（不再直接通过本地路径）。
```

### Docker部署

#### 单架构部署
```bash
# 构建镜像
make docker-build

# 运行容器
make docker-run

# 或使用完整部署命令
make deploy
```

#### 多架构部署
```bash
# 构建多架构镜像 (linux/amd64, linux/arm64)
make docker-buildx-local

# 部署多架构镜像
make deploy-multiarch

# 推送到镜像仓库 (需要先设置 REGISTRY 环境变量)
REGISTRY=your-registry.com make docker-buildx
```

#### 环境变量配置
```bash
# 设置镜像仓库地址
export REGISTRY=your-registry.com

# 设置镜像版本
export VERSION=1.0.0

# 构建并推送
make docker-buildx
```

### Kubernetes部署

#### 快速部署
```bash
# 进入部署目录
cd deployment

# 基础部署
./deploy.sh dev

# 完整部署 (包含MinIO、Ingress、HPA)
./deploy.sh dev --with-minio --with-ingress --with-hpa
```

#### 手动部署
```bash
# 创建命名空间
kubectl apply -f deployment/namespace.yaml

# 应用配置
kubectl apply -f deployment/configmap.yaml
kubectl apply -f deployment/secret.yaml

# 部署应用
kubectl apply -f deployment/deployment.yaml
kubectl apply -f deployment/service.yaml

# 可选组件
kubectl apply -f deployment/minio.yaml      # MinIO存储
kubectl apply -f deployment/ingress.yaml    # Ingress入口
kubectl apply -f deployment/hpa.yaml       # 自动扩缩容
```

#### 访问服务
```bash
# 集群内访问
http://ohc-account-invoice-service.ohc-account-invoice.svc.cluster.local:8000

# 端口转发访问
kubectl port-forward service/ohc-account-invoice-service 8000:8000 -n ohc-account-invoice

# 通过Ingress访问 (如果配置了Ingress)
http://ohc-account-invoice.local
```

#### 卸载
```bash
# 自动卸载
./deployment/undeploy.sh

# 手动卸载
kubectl delete namespace ohc-account-invoice
```

## API端点使用

### 可用的API端点

#### 通用端点
1. **GET /** - 根路径，返回API基本信息
2. **GET /health** - 健康检查
3. **GET /templates** - 获取所有支持的账票模板列表
4. **GET /templates/{name}** - 获取指定模板的详细信息
5. **POST /generate** - 生成账票文档（通用接口）
6. **GET /config** - 获取服务配置信息
7. **GET /download/{filename}** - 下载文件（本地存储时可用）

#### 专门的模板接口
每个模板都有专门的接口，提供更清晰的参数说明和验证：

1. **POST /generate/dhf-index** - 生成制作文档・图纸一览
2. **POST /generate/ptf-index** - 生成PTF INDEX
3. **POST /generate/es-individual-test-spec** - 生成ES个别试验要项书
4. **POST /generate/es-individual-test-result** - 生成ES个别试验结果书
5. **POST /generate/pp-individual-test-result** - 生成PP个别试验结果书
6. **POST /generate/es-verification-plan** - 生成ES验证计划书
7. **POST /generate/es-verification-result** - 生成ES验证结果书
8. **POST /generate/pp-verification-plan** - 生成PP验证计划书
9. **POST /generate/pp-verification-result** - 生成PP验证结果书
10. **POST /generate/basic-specification** - 生成基本规格书
11. **POST /generate/pp-individual-test-spec** - 生成PP个别试验要项书
12. **POST /generate/follow-up-dr-minutes** - 生成跟进DR会议记录
13. **POST /generate/labeling-specification** - 生成标签规格书
14. **POST /generate/product-environment-assessment** - 生成产品环境评估文档
15. **POST /generate/existing-product-comparison** - 生成与现有产品对比表

### 使用示例

#### Python客户端示例
```python
import requests
import json

# 1. 获取模板列表
response = requests.get("http://localhost:8000/templates")
templates = response.json()
print(f"支持的模板: {templates}")

# 2. 生成账票文档（通用接口）
generate_data = {
    "template_name": "DHF_INDEX",
    "parameters": {
        "project_name": "OHC项目",
        "date": "2025-01-22",
        "author": "张三",
        "version": "1.0",
        "document_type": "设计文档",
        "department": "研发部"
    }
}
response = requests.post("http://localhost:8000/generate", json=generate_data)
result = response.json()
print(f"生成结果: {result}")

# 3. 使用专门接口生成文档
dhf_data = {
    "project_name": "OHC项目",
    "date": "2025-01-22",
    "author": "张三",
    "version": "1.0",
    "document_type": "设计文档",
    "department": "研发部",
    "reviewer": "李四",
    "approval_date": "2025-01-23"
}
response = requests.post("http://localhost:8000/generate/dhf-index", json=dhf_data)
result = response.json()
print(f"生成结果: {result}")

# 4. 下载文件（本地存储）
if result.get('success') and result.get('file_name'):
    filename = result['file_name']
    download_response = requests.get(f"http://localhost:8000/download/{filename}")
    with open(filename, 'wb') as f:
        f.write(download_response.content)
    print(f"文件已下载: {filename}")
```

#### JavaScript客户端示例
```javascript
// 1. 获取模板列表
const templatesResponse = await fetch('http://localhost:8000/templates');
const templates = await templatesResponse.json();
console.log('支持的模板:', templates);

// 2. 生成账票文档（通用接口）
const generateData = {
    template_name: "DHF_INDEX",
    parameters: {
        project_name: "OHC项目",
        date: "2025-01-22",
        author: "张三",
        version: "1.0",
        document_type: "设计文档",
        department: "研发部"
    }
};
const generateResponse = await fetch('http://localhost:8000/generate', {
    method: 'POST',
    headers: {
        'Content-Type': 'application/json',
    },
    body: JSON.stringify(generateData)
});
const result = await generateResponse.json();
console.log('生成结果:', result);

// 3. 使用专门接口生成文档
const dhfData = {
    project_name: "OHC项目",
    date: "2025-01-22",
    author: "张三",
    version: "1.0",
    document_type: "设计文档",
    department: "研发部",
    reviewer: "李四",
    approval_date: "2025-01-23"
};
const dhfResponse = await fetch('http://localhost:8000/generate/dhf-index', {
    method: 'POST',
    headers: {
        'Content-Type': 'application/json',
    },
    body: JSON.stringify(dhfData)
});
const dhfResult = await dhfResponse.json();
console.log('生成结果:', dhfResult);
```

## 默认值功能

### 自动默认值

为了简化API使用，系统为常用字段提供了智能默认值：

- **date**: 自动设置为当前服务器时间（精确到秒）
- **author**: 默认设置为 "OHC账票AI助手"

### 使用默认值

```bash
# 最小请求（使用默认值）
curl -X POST "http://localhost:8000/generate/dhf-index" \
  -H "Content-Type: application/json" \
  -d '{
    "project_name": "OHC项目",
    "version": "1.0",
    "document_type": "设计文档",
    "department": "研发部"
  }'
```

### 覆盖默认值

```bash
# 完整请求（覆盖默认值）
curl -X POST "http://localhost:8000/generate/dhf-index" \
  -H "Content-Type: application/json" \
  -d '{
    "project_name": "OHC项目",
    "version": "1.0",
    "date": "2025-01-22 15:30:45",
    "author": "张三",
    "document_type": "设计文档",
    "department": "研发部",
    "reviewer": "李四",
    "approval_date": "2025-01-23"
  }'
```

### 支持的日期格式

- `YYYY-MM-DD HH:MM:SS` (2025-01-22 15:30:45)
- `YYYY-MM-DD` (2025-01-22)
- `YYYY/MM/DD` (2025/01/22)
- `YYYYMMDD` (20250122)
- `DD/MM/YYYY` (22/01/2025)
- `MM/DD/YYYY` (01/22/2025)
- `YYYY-MM-DD HH:MM` (2025-01-22 15:30)

## 高级功能

### 智能模板填充策略

服务采用策略模式设计，每个模板都有专门的填充策略，并支持智能文件类型识别：

#### 核心优化特性

- **智能文件类型识别**: 根据模板文件后缀自动选择Excel或Word填充方式
- **跨格式填充支持**: 每个填充策略都能处理多种文件类型
- **精确模板匹配**: 根据输出文件格式选择对应的模板文件
- **增强错误处理**: 提供更精确的错误信息和调试信息
- **优化占位符替换**: 使用正则表达式直接替换整个占位符（包括括号）
- **智能文件名生成**: 按照"项目号_版本号_模版文件名_日期时间数字"格式生成文件名

#### 专用填充策略

- **DHFIndexFiller**: DHF INDEX专用填充策略
  - 支持Excel和Word文件智能识别
  - 项目信息填充和文档列表管理
  - 根据文件类型选择专用填充方法

- **PTFIndexFiller**: PTF INDEX专用填充策略
  - 支持Excel和Word文件智能识别
  - 测试信息填充和测试阶段管理
  - 智能测试环境配置

- **TestSpecFiller**: 试验要项书专用填充策略
  - 支持Excel和Word文件智能识别
  - 试验规格表格和试验要求填充
  - 智能试验参数配置

- **VerificationPlanFiller**: 验证计划书专用填充策略
  - 支持Excel和Word文件智能识别
  - 验证计划表格和验证范围填充
  - 智能验证流程管理

- **MeetingMinutesFiller**: 会议记录专用填充策略
  - 支持Excel和Word文件智能识别
  - 会议信息填充和会议结构管理
  - 智能会议内容组织

- **SmartWordFiller**: 智能Word填充策略
  - 支持Excel和Word文件智能识别
  - 智能内容填充、智能表格处理
  - 结构化文档处理

- **AdvancedExcelFiller**: 高级Excel填充策略
  - 支持Excel和Word文件智能识别
  - 动态表格、图表数据填充
  - 列表数据处理

### 占位符替换优化

#### 支持的占位符格式

- **单括号格式**: `{变量名}` - 标准占位符格式
- **双括号格式**: `{{变量名}}` - 增强占位符格式，优先级更高

#### 替换特性

- **完整替换**: 使用正则表达式直接替换整个占位符（包括括号）
- **优先级处理**: 双括号格式优先于单括号格式，避免嵌套问题
- **智能处理**: 不存在的参数保持原样，不进行替换
- **高性能**: 使用正则表达式替换，性能更优
- **复杂文本支持**: 支持包含特殊字符和复杂结构的文本

#### 使用示例

```text
原始文本: "项目名称：{project_name}，版本：{{version}}，日期：{date}"
替换后: "项目名称：OHC项目，版本：1.0，日期：2025-01-22"

原始文本: "项目：{project_name}（版本{version}）"
替换后: "项目：OHC项目（版本1.0）"

原始文本: "项目：{{project_name}}，不存在的参数：{{non_existent}}"
替换后: "项目：OHC项目，不存在的参数：{{non_existent}}"
```

### 智能文件名生成

#### 文件名格式

生成的文件名遵循统一格式：`项目号_版本号_模版文件名_日期时间数字.扩展名`

#### 格式说明

- **项目号**: 从参数中提取的项目名称，支持中文字符
- **版本号**: 从参数中提取的版本信息
- **模版文件名**: 模板的显示名称，支持中文和日文字符
- **日期时间数字**: 格式为 `YYYYMMDD_HHMMSS`
- **扩展名**: 根据模板类型自动选择（.xlsx 或 .docx）

#### 字符处理

- **保留字符**: 中文字符、日文字符、英文字母、数字、连字符、下划线、点号
- **特殊字符**: 其他特殊字符会被替换为下划线
- **默认值**: 缺少必要参数时提供合理的默认值

#### 使用示例

```text
输入参数:
{
  "project_name": "OHC测试项目",
  "version": "1.0",
  "date": "2025-01-22 15:30:45"
}

生成文件名: OHC测试项目_1.0_ドキュメント・図面一覧_20250122_153045.xlsx

文件名结构:
- 项目号: OHC测试项目
- 版本号: 1.0
- 模版名: ドキュメント・図面一覧
- 日期时间: 20250122_153045
- 扩展名: .xlsx
```

#### 支持的日期格式

- `YYYY-MM-DD HH:MM:SS` (2025-01-22 15:30:45)
- `YYYY-MM-DD HH:MM` (2025-01-22 15:30)
- `YYYY-MM-DD` (2025-01-22)
- `YYYY/MM/DD` (2025/01/22)
- `YYYYMMDD` (20250122)
- `DD/MM/YYYY` (22/01/2025)
- `MM/DD/YYYY` (01/22/2025)

### 智能内容识别

智能Word填充策略会根据内容类型自动选择填充方式：

- **项目概述**: 自动识别并填充项目概述内容
- **技术要求**: 智能填充技术要求段落
- **验收标准**: 自动填充验收标准内容
- **表格数据**: 根据表格结构智能填充数据

### 动态表格填充

支持在Excel模板中填充动态表格数据：

```json
{
  "product_list": [
    {"name": "产品A", "price": "1000", "features": "基础功能"},
    {"name": "产品B", "price": "1500", "features": "高级功能"},
    {"name": "产品C", "price": "2000", "features": "专业功能"}
  ]
}
```

### 图表数据填充

支持填充图表数据：

```json
{
  "chart_data": {
    "performance_chart": [85, 92, 78, 96, 88],
    "cost_chart": [1000, 1200, 1100, 1300, 1250]
  }
}
```

## 配置说明

### 环境变量

```bash
# 存储配置（固定为 MinIO）
STORAGE_TYPE=minio

# MinIO 配置（必填）
MINIO_ENDPOINT=localhost:9000
MINIO_ACCESS_KEY=minioadmin
MINIO_SECRET_KEY=minioadmin
MINIO_BUCKET_NAME=ohc-documents
MINIO_SECURE=false

# 应用配置
APP_NAME=OHC账票生成服务
APP_VERSION=1.0.0
HOST=0.0.0.0
PORT=8000
DEBUG=false
RELOAD=false

# 模板配置
TEMPLATE_BASE_PATH=static/templates

# 文件名配置
FILENAME_INCLUDE_TIMESTAMP=true
FILENAME_MAX_LENGTH=200
```

### 在 CI / 非完整基础设施环境下运行

在某些 CI 或受限环境中，系统可能无法访问本地配置文件或外部依赖（如 MinIO）。为保证自动化生成 OpenAPI、运行单元测试或收集代码静态信息时不会因外部服务未准备好而失败，项目支持以下环境变量开关：

- `SKIP_INFRA_INIT=1`  
  - 含义：跳过基础设施（MinIO 等）初始化，服务将不会在导入时尝试连接或实例化外部存储客户端。适用于 CI、静态分析或仅运行单元测试的场景。  
  - 在 CI 中我们已将该变量设置在 `CI` workflow，以确保 lint/tests 在没有 MinIO 的环境下可以通过。

- `CONFIG_FILE=/path/to/config.toml`  
  - 含义：指定自定义配置文件路径，用于覆盖默认的配置文件查找逻辑。如果不设置，系统会自动查找 `config/config.toml` 或 `config.toml`。

示例（本地生成 OpenAPI 时推荐）：

```bash
# 跳过 infra 初始化并生成 OpenAPI 文件（结果写入 src/swagger/）
SKIP_INFRA_INIT=1 python tools/generate_openapi.py
```

生成的文件位置：
- `src/swagger/openapi.json` — OpenAPI JSON 描述文件  
- `src/swagger/swagger.html` — 静态 Swagger UI，可直接在浏览器打开（相对路径会加载 `openapi.json`）
 
服务生成的输出文件（当使用本地存储时）位于：`<LOCAL_STORAGE_PATH>/{project}/{version}/...`，LOCAL_STORAGE_PATH 默认为 `generated_files`（可通过环境变量覆盖）。

CI 注意事项：
- GitHub Actions workflow 已在 `.github/workflows/generate-openapi.yml` 中生成并上传 `src/swagger` 作为 artifact。  
- 若在 CI 中需要运行与 MinIO 交互的集成测试，请在 workflow 中提供相应的 MinIO 服务或取消 `SKIP_INFRA_INIT` 并注入正确的环境变量。

## 开发指南

### 项目结构（精简后）

```
ohc_account_invoice/
├── src/
│   ├── main.py                 # FastAPI 应用主文件
│   ├── config.py               # 配置管理（包含 STORAGE_TYPE 等）
│   ├── application/            # 业务用例层（application）
│   ├── infrastructure/         # 基础设施层（存储、模板实现）
│   ├── interfaces/             # HTTP 路由与 Pydantic 模型
│   ├── swagger/                # 生成的 OpenAPI/Swagger 静态文件
├── deployment/                 # Kubernetes 与部署脚本
├── tests/                      # 测试文件
├── pyproject.toml              # 项目配置（运行时依赖）
├── Makefile                    # 构建脚本
├── Dockerfile                  # Docker 配置
└── README.md                   # 项目文档
```

### 添加新模板

1. 在 `src/templates/` 目录下添加模板文件
2. 在 `src/infrastructure/template_service.py` 中添加模板实现或策略
3. 在 `src/interfaces/schemas.py`（或 `src/interfaces/schemas/templates.py`）中添加 API 请求/响应的 Pydantic 模型
4. 在 `src/application/` 中实现用例（application 层），并在 `src/main.py` 中通过路由暴露端点

### 自定义填充策略

```python
class CustomFiller(TemplateFillerStrategy):
    """自定义填充策略"""
    
    def fill_template(self, template_path: Path, parameters: Dict[str, Any], output_path: Path) -> bool:
        # 实现自定义填充逻辑
        pass
```

## 故障排除

### 常见问题

1. **模板文件不存在**
   - 检查 `src/templates/` 目录下是否有对应的模板文件
   - 确认文件名和模板名称匹配

2. **参数验证失败**
   - 检查参数格式和类型
   - 参考API文档中的参数说明

3. **文件生成失败**
   - 检查模板文件是否损坏
   - 确认参数是否完整

4. **存储配置错误**
   - 检查环境变量配置
   - 确认存储路径权限

### 日志调试

```bash
# 启用调试模式
export DEBUG=true
make dev
```

## 许可证

MIT License

## 贡献

欢迎提交Issue和Pull Request！

## 更新日志

### v1.0.0 (2025-01-22)
- 初始版本发布
- 支持15种账票模板
- 基于FastAPI框架
- 支持 MinIO 存储
- 策略模式模板填充
- 智能内容识别
- 完整的API文档
