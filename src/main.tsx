import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import './styles.css';
import { Capacitor } from '@capacitor/core';
import { Directory, Filesystem } from '@capacitor/filesystem';
import { Share } from '@capacitor/share';

// Android WebView does not consistently handle Blob URL downloads.
// Intercept only native APK downloads; normal browser behavior stays untouched.
if (Capacitor.isNativePlatform() && typeof HTMLAnchorElement !== 'undefined') {
  const originalAnchorClick = HTMLAnchorElement.prototype.click;
  HTMLAnchorElement.prototype.click = function () {
    const anchor = this;
    if (!anchor.hasAttribute('download') || !anchor.href.startsWith('blob:')) {
      return originalAnchorClick.call(anchor);
    }
    void (async () => {
      const response = await fetch(anchor.href);
      if (!response.ok) throw new Error('Não foi possível preparar o arquivo.');
      const blob = await response.blob();
      const dataUrl = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result || ''));
        reader.onerror = () => reject(reader.error);
        reader.readAsDataURL(blob);
      });
      const base64 = dataUrl.split(',')[1] || '';
      const fileName = (anchor.download || 'coletaflow-arquivo').replace(/[\\/:*?"<>|]/g, '-');
      const saved = await Filesystem.writeFile({
        path: fileName,
        data: base64,
        directory: Directory.Documents,
        recursive: true
      });
      await Share.share({
        title: fileName,
        text: 'Arquivo exportado pelo ColetaFlow',
        url: saved.uri,
        dialogTitle: 'Arquivo pronto para salvar ou compartilhar'
      });
      URL.revokeObjectURL(anchor.href);
    })().catch(() => originalAnchorClick.call(anchor));
  };
}

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
