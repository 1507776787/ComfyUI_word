# ComfyUI_word 节点
> 基于 HTML 标签结构自动生成 PPT 文件的 ComfyUI 自定义节点

## 节点名称
`PPT`

## 功能介绍
- 接收 `positive` 端口输入的纯 HTML 幻灯片文本，自动生成标准 PPT 文件
- 固定画幅比例：**12.8:7.2**，内置字体与排版约束，防止文字出框
- 支持多页幻灯片自动排版，长文本精简拆页，每页只保留一个重点
- 内置 `orange_glass` 主题，支持浅底深字 / 深底浅字配色
- 支持卡片、列表、表格、图表、目录、按钮、标注等富组件渲染
- 遵循严格布局规范：双栏布局、尺寸间距、卡片比例统一约束
- 最终生成可直接下载的 PPT 文件并返回下载链接

## 安装方法
1. 进入 ComfyUI 根目录下的 `custom_nodes` 文件夹
2. 克隆本仓库：
```bash
git clone https://github.com/1507776787/ComfyUI_word.git
```
3. 重启 ComfyUI，节点自动加载

## 输入 / 输出
### 输入
- `positive`：以 `<slide>` 开始、`</slide>` 结束的纯 HTML 幻灯片文本

### 输出
- 生成完整 PPT 文件，输出下载链接 / 保存路径

## 使用规范
1. **支持标签**
slide、theme、layout、section、background、footer、toc、spacer、divider、title、h1、h2、p、ul、ol、li、card、badge、button、icon、callout、chart、table

2. **字体大小约束**
- title：30~36px
- h2：18~22px
- 正文 / 列表：13~15px

3. **卡片规则**
card 必须使用属性式写法：
```html
<card title="..." body="..." />
```

4. **布局限制**
- 每页最多 1 个主标题、2 个 section
- 每个 section 最多 3 张 card 或 5 条 li
- 双栏布局：左栏 x="1.0" width="5.1"；右栏 x="6.3" width="5.1"

5. **内容建议**
- 单次生成 5~6 页为宜
- 文字过长自动拆页，避免拥挤溢出

## 示例代码
```html
<theme name="orange_glass" />
<footer text="演示文稿 | 第 {page}/{total} 页 | {date}" />

<slide>
  <layout name="default" />
  <background color="#FFF7ED" />
  <title style="font-size:34;color:#7C2D12;">PPT 自动生成演示</title>
  <section x="1.4" y="2.0" width="10.0" height="3.9" bg="#FFFFFFCC" padding="0.24">
    <h2 style="font-size:21;color:#7C2D12">HTML 驱动智能排版</h2>
    <p style="font-size:15;color:#9A3412">多组件支持，自动防文字溢出</p>
  </section>
</slide>
```

## 更新日志
- v1.0 初始版本
  - 支持 HTML 结构转 PPT 完整功能
  - 支持主题、双栏、卡片、表格、图表等组件
  - 内置防文字出框排版规则
