// Read engine.js at build time and export as a string
// Vite will inline this via ?raw import
export { default as workerScript } from './engine.js?raw';
