# EC Mall Skills

这个仓库目前提供一个 Codex skill：`amazon-listing-xlsx`。

它用于把固定格式的商品信息 Excel 文件转换成固定格式的 Amazon listing Excel 文件。说白了，就是让用户丢一个商品源表，agent 生成标题、描述、搜索词和五点描述，再输出可交付的 `.xlsx`。别让人手搓表格了，太原始，像拿算盘跑云计算。

## Skill 作用

`amazon-listing-xlsx` 适用于以下场景：

- 输入文件是固定格式的商品信息 `.xlsx`
- 源表 sheet 名为 `Product Info`
- 每行代表一个商品
- 每行必须有非空 `Item Code`
- 输出文件需要对齐固定 Amazon listing 模板

输出 workbook 的结构固定：

- sheet 名：`listings`
- 列顺序：
  - `itemCode`
  - `productTitle`
  - `productDescription`
  - `searchTerms`
  - `bulletPoint1`
  - `bulletPoint2`
  - `bulletPoint3`
  - `bulletPoint4`
  - `bulletPoint5`

skill 内置了：

- `SKILL.md`：agent 使用入口
- `references/listing_prompt_rules.md`：Amazon listing 文案规则
- `scripts/listing_workbook.py`：Excel 解析、写入和校验脚本

## 安装方式

### 推荐：让 agent 安装

对 Codex 或支持 skills 的 agent 说：

```text
$skill-installer
请从 GitHub 安装这个 skill：
repo: VincentChris/ec-mall-skills
path: amazon-listing-xlsx
```

安装后重启 Codex，让新 skill 生效。

### 手动安装

如果你要手动安装到 Codex skills 目录：

```bash
git clone https://github.com/VincentChris/ec-mall-skills.git
mkdir -p "${CODEX_HOME:-$HOME/.codex}/skills"
cp -R ec-mall-skills/amazon-listing-xlsx "${CODEX_HOME:-$HOME/.codex}/skills/"
```

然后重启 Codex。

## 使用方式

安装后，直接告诉 agent：

```text
使用 amazon-listing-xlsx skill，把 /path/to/source.xlsx 生成 Amazon listing xlsx。
```

agent 会按 skill 流程执行：

1. 读取商品源文件
2. 导出商品行 JSON
3. 根据 listing 文案规则生成标题、描述、搜索词和五点描述
4. 写入固定格式的 Amazon listing `.xlsx`
5. 校验输出文件
6. 返回生成的 `.xlsx` 路径

## 脚本直接用法

如果你只想直接跑 workbook 工具，可以设置 skill 目录：

```bash
SKILL_DIR=/path/to/amazon-listing-xlsx
```

导出源数据：

```bash
python "$SKILL_DIR/scripts/listing_workbook.py" export-source source.xlsx --output build/source.rows.json
```

把生成好的 listing JSON 写成 Excel：

```bash
python "$SKILL_DIR/scripts/listing_workbook.py" write-output source.xlsx build/generated-listings.json --output source.amazon-listings.xlsx
```

校验输出：

```bash
python "$SKILL_DIR/scripts/listing_workbook.py" validate-output source.xlsx source.amazon-listings.xlsx
```

注意：`generated-listings.json` 是中间产物，不是最终交付物。最终交付物是 `.xlsx`。

## 开发和测试

安装测试依赖：

```bash
python3 -m pip install -r requirements-dev.txt
```

运行测试：

```bash
pytest -q
```

当前测试覆盖：

- 源表解析
- 必填表头校验
- 缺失 `Item Code` 拒绝
- 输出表头和 sheet 校验
- 行数匹配校验
- 重复 `rowIndex` 拒绝
- 空目标字段拒绝
- 写入失败不留下坏文件

## 示例文件

仓库根目录包含：

- `source.xlsx`：商品信息源文件示例
- `target.xlsx`：Amazon listing 目标格式示例

这两个文件用于理解固定输入/输出格式和做端到端验证。
