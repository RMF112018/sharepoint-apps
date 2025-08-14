require('@testing-library/jest-dom');

// Register fluent icons to silence icon warnings in tests
const { initializeIcons } = require('@fluentui/font-icons-mdl2');
try { initializeIcons(); } catch {}

