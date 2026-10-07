import { defineConfig, type Plugin } from 'vite';

// A self-contained distribution that works with file://, without CDNs or a server.
function portableHTML(): Plugin {
  return {
    name: 'folha-portable-html',
    enforce: 'post',
    generateBundle: { order: 'post', handler(_options, bundle) {
      const html = bundle['index.html'];
      if (!html || html.type !== 'asset') return;
      let content = String(html.source);
      for (const [name, output] of Object.entries(bundle)) {
        if (output.type === 'chunk') {
          content = content.replace(new RegExp(`<script[^>]*src="[^\"]*${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*><\\/script>`),
            () => `<script type="module">${output.code.replace(/<\/script/gi, '<\\/script')}</script>`);
          delete bundle[name];
        } else if (name.endsWith('.css')) {
          content = content.replace(/<link[^>]*rel="stylesheet"[^>]*>/, () => `<style>${output.source}</style>`);
          delete bundle[name];
        }
      }
      html.source = content;
    } },
  };
}

export default defineConfig(({ mode, command }) => ({
  base: mode === 'portable' ? './' : command === 'build' ? '/folha/' : '/',
  plugins: mode === 'portable' ? [portableHTML()] : [],
  build: {
    chunkSizeWarningLimit: 1000,
    ...(mode === 'portable' ? {
      cssCodeSplit: false,
      modulePreload: false,
      assetsInlineLimit: Infinity,
      rollupOptions: { output: { inlineDynamicImports: true } },
    } : {}),
  },
}));
