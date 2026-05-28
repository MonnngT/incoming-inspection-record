# 来料检验记录系统（跳批检验逻辑）

一个基于 **Streamlit + Google Sheets** 的 IQC 来料检验记录小程序，自动根据跳批检验规则判定"正常检验 / 跳批"，并记录检验用时。

---

## ✨ 功能特性

- **到货日期**：日历下拉选择
- **供应商**：下拉选择（速锐达(SP)、AQ、压铸件）
- **零件料号**：根据供应商联动下拉（压铸件会显示子供应商如 新*/金*/欧* 便于区分）
- **生产日期**：日历下拉，或从该料号历史生产日期中选择
- **累计批次数**：根据历史记录自动计算（同一料号+同一生产日期）
- **执行动作**：根据跳批规则自动判定
  - `正常检验`
  - `跳批（不检验尺寸）` — 跳过的批次
  - `跳批检验（全项目）` — 跳批序列里需要检的那一批
- **检验用时**：分"开始时间 / 结束时间"两个输入，自动算出用时（分钟）
- **结果**：下拉 OK / NG
- **检验员**：下拉 杨明 / 田志高 / 其他（选"其他"可手动输入姓名）
- **历史记录**：全部存入 Google Sheets，可筛选、统计、导出 CSV

---

## 🔁 跳批规则逻辑（已内置）

1. 跳批序列以 **"料号 + 生产日期"** 为单位累计（同一生产日期才能跳批）
2. 连续 **3 批合格** → 启动跳批，进入 **"跳2检1"** 循环（跳过2批、检验1批）
3. 跳批中 **任一批不合格** → 退回正常检验，连续合格计数清零，重新累计3批
4. 跳批仅免做 **尺寸** 检验；外观、包装数量每批必检

---

## 🚀 部署步骤（零基础也能照做）

### 第一步：准备 Google Sheet

1. 用浏览器打开 [Google Sheets](https://sheets.google.com)，新建一个空白表格
2. 给表格取个名字，比如"来料检验记录"
3. 看浏览器地址栏：
   `https://docs.google.com/spreadsheets/d/`**`1AbCdEfGhIjK...`**`/edit`
   中间加粗那一段就是 **Sheet Key**，先复制保存下来

### 第二步：创建 Google 服务账号（让程序能读写表格）

1. 打开 [Google Cloud Console](https://console.cloud.google.com)
2. 新建一个项目（Project）
3. 左侧菜单 → "API和服务" → "已启用的API和服务" → 点"启用API和服务"
   - 搜索并启用 **Google Sheets API**
   - 搜索并启用 **Google Drive API**
4. 左侧菜单 → "凭据" → "创建凭据" → "服务账号"
   - 填个名字（如 inspection-bot）→ 创建并继续 → 完成
5. 点进刚创建的服务账号 → "密钥" 选项卡 → "添加密钥" → "创建新密钥" → 选 **JSON** → 下载
   - 会下载一个 `xxxxx.json` 文件，**这是密钥，保管好，不要泄露**
6. 打开这个 JSON 文件，找到 `client_email` 字段，复制那个邮箱地址
   （形如 `inspection-bot@xxx.iam.gserviceaccount.com`）
7. 回到第一步建的 Google Sheet → 点右上角"共享" → 把上面那个邮箱加进去 → 权限设为 **编辑者**

### 第三步：上传代码到 GitHub

1. 在 [GitHub](https://github.com) 新建一个仓库（Repository），比如 `iqc-inspection`
2. 把以下文件传上去：
   - `app.py`
   - `parts_data.json`
   - `requirements.txt`
   - `.gitignore`
   - `README.md`
   - **不要上传** `secrets.toml`（含密钥）！.gitignore 已帮你屏蔽

### 第四步：部署到 Streamlit Cloud

1. 打开 [share.streamlit.io](https://share.streamlit.io)，用 GitHub 账号登录
2. 点 "New app" → 选择你的仓库 → main 分支 → 主文件填 `app.py` → Deploy
3. 部署后点右下角 "⚙️ Settings" → "Secrets"，粘贴以下内容（按 `secrets.toml.example` 填真实值）：

```toml
[sheet]
key = "你的Google Sheet Key"

[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "inspection-bot@xxx.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
universe_domain = "googleapis.com"
```

> 💡 把下载的 JSON 文件内容对照着填进去即可。`private_key` 那一长串要保留 `\n`。

4. 保存后 App 会自动重启，就能用了！

---

## 🖥️ 本地测试（可选）

```bash
pip install -r requirements.txt
mkdir .streamlit
cp secrets.toml.example .streamlit/secrets.toml
# 编辑 .streamlit/secrets.toml 填入真实密钥
streamlit run app.py
```

浏览器会自动打开 `http://localhost:8501`

---

## 📁 文件说明

| 文件 | 作用 |
|------|------|
| `app.py` | 主程序 |
| `parts_data.json` | 供应商-料号映射数据（从你的 Excel 提取） |
| `requirements.txt` | Python 依赖 |
| `secrets.toml.example` | 密钥配置模板 |
| `.gitignore` | 防止密钥误传 GitHub |

---

## 🔧 如何修改料号 / 供应商 / 检验员

- **改料号**：编辑 `parts_data.json`，按现有格式增删即可
- **改检验员名单**：编辑 `app.py` 里的 `INSPECTORS = ["杨明", "田志高", "其他"]`
- **改跳批规则**（如连续几批合格才跳批、跳几检几）：编辑 `app.py` 顶部的
  `CONSECUTIVE_PASS_TO_START`、`SKIP_PATTERN_SKIP`、`SKIP_PATTERN_INSPECT`
