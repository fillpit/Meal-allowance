const fs = require('fs');
const path = require('path');

try {
  console.log('开始打包 食堂补贴数据生成器.html...');
  
  // 读取源文件
  let html = fs.readFileSync(path.join(__dirname, 'index.html'), 'utf8');
  const css = fs.readFileSync(path.join(__dirname, 'style.css'), 'utf8');
  const js = fs.readFileSync(path.join(__dirname, 'app.js'), 'utf8');
  
  // 替换 CSS 引用为内嵌 style 标签
  const cssRegex = /<link\s+rel="stylesheet"\s+href="style\.css"\s*\/?>/i;
  if (cssRegex.test(html)) {
    html = html.replace(cssRegex, `<style>\n${css}\n</style>`);
    console.log('✓ 已成功内嵌 style.css');
  } else {
    console.warn('⚠️ 未在 index.html 中找到 style.css 引用，请检查文件名或路径');
  }

  // 检测并内嵌 avatar.png 作为 Base64 数据 URL（若本地存在该文件）
  const avatarPath = path.join(__dirname, 'avatar.png');
  const avatarRegex = /src="avatar\.png"/gi;
  if (fs.existsSync(avatarPath)) {
    const avatarBase64 = fs.readFileSync(avatarPath).toString('base64');
    html = html.replace(avatarRegex, `src="data:image/png;base64,${avatarBase64}"`);
    console.log('✓ 已将 avatar.png 转换为 Base64 嵌入 HTML 交付包中');
  } else {
    console.log('💡 未在本地找到 avatar.png，交付包在没有该文件时将默认自动降级使用精美的矢量青蛙卡通头像');
  }
  
  // 替换 JS 引用为内嵌 script 标签
  const jsRegex = /<script\s+src="app\.js"\s*><\/script>/i;
  if (jsRegex.test(html)) {
    html = html.replace(jsRegex, `<script>\n${js}\n</script>`);
    console.log('✓ 已成功内嵌 app.js');
  } else {
    console.warn('⚠️ 未在 index.html 中找到 app.js 引用，请检查文件名或路径');
  }
  
  // 写入最终文件
  const outputPath = path.join(__dirname, '食堂补贴数据生成器.html');
  fs.writeFileSync(outputPath, html, 'utf8');
  console.log(`\n🎉 打包成功！输出文件: ${outputPath}`);
  console.log('现在您可以直接双击运行该 HTML 文件了！');
} catch (error) {
  console.error('❌ 打包出错:', error.message);
  process.exit(1);
}
