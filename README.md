# 亚马逊品牌旗舰店竞品调研 Excel 生成工具

本仓库提供 `generate_amazon_store_research.py`，用于生成“亚马逊品牌旗舰店竞品调研”Excel 模板。

> Pull Request 不直接提交 `.xlsx` 二进制文件。需要 Excel 时，请通过 GitHub Actions 手动生成并下载。

## 如何在 GitHub Actions 中生成 Excel

1. 进入当前 GitHub 仓库页面。
2. 点击顶部导航栏中的 **Actions**。
3. 在左侧工作流列表中选择 **生成亚马逊旗舰店调研Excel**。
4. 点击 **Run workflow**。
5. 在弹出的确认区域继续点击 **Run workflow** 开始运行。
6. 等待工作流运行完成。
7. 打开本次运行记录，在页面底部的 **Artifacts** 区域下载 `amazon-store-research-excel`。
8. 下载并解压 artifact 后，即可获得 `亚马逊品牌旗舰店竞品调研.xlsx`。

## 本地生成方式

如果需要在本地生成 Excel，可执行：

```bash
python -m pip install -r requirements.txt
python generate_amazon_store_research.py
```

运行成功后，仓库根目录会生成：

```text
亚马逊品牌旗舰店竞品调研.xlsx
```

## 生成内容

脚本会生成以下 5 个工作表：

- `02 竞品清单`
- `03 店铺调研`
- `04 优秀案例`
- `05 调研结论`
- `06 装修建议`
