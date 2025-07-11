# Excel列复制工具

这是一个简单的Web应用程序，允许用户选择两个Excel文件，并将源Excel文件中的指定列复制到目标Excel文件中的指定列。

## 功能

- 上传源Excel文件和目标Excel文件
- 选择源文件和目标文件中的工作表
- 选择源列和目标列
- 将源列数据复制到目标列
- 下载处理后的Excel文件

## 技术栈

- Next.js
- React
- TypeScript
- TailwindCSS
- xlsx库（用于Excel文件处理）
- react-dropzone（用于文件上传）

## 本地开发

1. 克隆仓库
2. 安装依赖：`npm install`
3. 启动开发服务器：`npm run dev`
4. 在浏览器中打开：`http://localhost:3000`

## 部署

此项目已配置为使用Vercel进行部署。只需将代码推送到GitHub仓库，然后在Vercel上导入该仓库即可。 