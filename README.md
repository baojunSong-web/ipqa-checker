# 自工程检查重点线体工具

## 功能
- 读取生产计划PPT文件
- 根据过去40天数据分析各线体主力产品
- 标记产品变化线体为重点检查对象
- 支持导出PDF、发送邮件

## GitHub Actions 云打包说明

### 首次配置步骤

1. **创建GitHub仓库**
   - 访问 https://github.com/new
   - 仓库名称：`ipqa-checker`（私有）
   - 不要勾选任何初始化选项
   - 点击 Create repository

2. **上传代码**
   ```bash
   cd /home/song/.openclaw/workspace/ipqa-checker
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/你的用户名/ipqa-checker.git
   git push -u origin main
   ```

3. **触发打包**
   - 访问仓库页面 → Actions
   - 点击 "Build EXE" 工作流
   - 点击 "Run workflow"
   - 等待3-5分钟

4. **下载EXE**
   - 点击 workflow 运行记录
   - 点击 "ipqa-checker-exe" artifacts
   - 下载并解压

## 依赖
- PyQt5
- python-pptx
- reportlab
- pyinstaller

## 使用说明
运行后：
1. 选择包含PPT文件的文件夹
2. 选择要分析的日期文件
3. 点击"分析变化点"
4. 可导出PDF或发送邮件
