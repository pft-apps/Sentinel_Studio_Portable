import './index.css';
import { startIntelStudio } from './intelEngine.ts';

if (document.readyState === 'loading') {
  window.addEventListener('DOMContentLoaded', () => {
    startIntelStudio();
  });
} else {
  startIntelStudio();
}
