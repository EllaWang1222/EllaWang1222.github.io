# EllaWang1222.github.io — 王佳辰的双语个人主页

这是一个由 Excel 驱动的静态个人网站，可部署到 GitHub Pages。

## 更新网站内容

1. 编辑 `content/个人简历.xlsx`。如果以后再次改名，只要 `content` 文件夹中只有一个 `.xlsx` 文件，脚本也会自动识别。
2. 每个工作表对应一个页面，工作表名称就是导航名称；第一个工作表是首页。
3. 每一列对应一个模块，第一行是模块标题，下面每个非空单元格是一条内容。
4. 图片单独占一个单元格，填写文件名（例如 `photo.jpg`），并把图片放入 `content/assets/`。
5. 需要中英文内容时，在同一个单元格中使用 `中文内容 || English content`。只写一种语言时，两种模式都会显示这段内容。
6. 运行 `python scripts/build_content.py`，再提交并推送到 GitHub。GitHub Actions 会自动发布更新。

## 本地预览

在项目目录运行：

```powershell
python scripts/build_content.py
python -m http.server 8000 --directory dist
```

然后打开 `http://localhost:8000`。

## GitHub Pages

仓库包含自动发布工作流。首次上传后，在仓库的 **Settings → Pages → Source** 中选择 **GitHub Actions**，之后每次推送都会更新网站。
