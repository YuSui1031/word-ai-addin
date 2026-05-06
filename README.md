# AI 智能续写 Word 插件

## 功能
- 自动检测输入文字，调用 AI 生成续写建议
- 灰色斜体显示在光标位置
- 按 `Ctrl+Tab` 采纳建议
- 继续输入则自动重新生成

## 安装步骤

### 1. 安装依赖
```
cd word-ai-addin
npm install
```

### 2. 启动服务器
```
npm start
```
服务器运行在 `https://localhost:3000`

### 3. 在 Word 中加载插件

**方法一：通过 Word 开发者工具**
1. 打开 Word → 文件 → 选项 → 自定义功能区
2. 勾选「开发工具」
3. 开发工具选项卡 → 加载项 → 共享文件夹
4. 将 `manifest.xml` 复制到共享文件夹路径下
5. 重启 Word，插件出现在加载项列表

**方法二：旁加载（推荐）**
1. 下载 Office-Addin-CLI
   ```
   npm install -g office-addin-dev-settings
   office-addin-dev-settings sideload manifest.xml
   ```

### 4. 配置 API
1. 打开 Word，点击「AI 续写」→「开启续写」
2. 在侧边栏填写 API 地址、Key、模型名称
3. 点击保存设置

## 支持的 API
- OpenAI API (api.openai.com)
- 任何兼容 OpenAI 格式的 API
- 国内可用: 智谱 GLM、通义千问（OpenAI 兼容模式）

## 注意事项
- 首次启动需允许自签名证书
- 需要稳定的网络连接
- 生成速度取决于 API 响应时间
