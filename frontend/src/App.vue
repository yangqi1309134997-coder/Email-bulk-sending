<template>
  <el-config-provider :locale="zhCn">
    <router-view />
  </el-config-provider>
</template>

<script setup>
import zhCn from 'element-plus/es/locale/lang/zh-cn'
import { ElConfigProvider } from 'element-plus'
</script>

<style>
:root {
  --app-bg: #f4f6fb;
  --app-surface: #ffffff;
  --app-border: #e6ebf2;
  --app-text: #1f2937;
  --app-muted: #6b7280;
  --app-primary: #2563eb;
  --app-primary-soft: #dbeafe;
  --app-success: #059669;
  --app-warning: #d97706;
  --app-danger: #dc2626;
  --app-shadow: 0 10px 30px rgba(15, 23, 42, 0.06);
  --app-radius: 14px;
}

* { box-sizing: border-box; }

html, body, #app {
  margin: 0;
  padding: 0;
  min-height: 100%;
  background: var(--app-bg);
  color: var(--app-text);
  font-family: "Segoe UI", "PingFang SC", "Microsoft YaHei", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
}

a { color: inherit; text-decoration: none; }

.el-card {
  border: 1px solid var(--app-border) !important;
  border-radius: var(--app-radius) !important;
  box-shadow: none !important;
  background: var(--app-surface);
}

.el-card__header {
  border-bottom: 1px solid var(--app-border) !important;
  font-weight: 600;
}

.el-button {
  border-radius: 10px !important;
}

.el-table {
  --el-table-header-bg-color: #f8fafc;
  --el-table-row-hover-bg-color: #f5f8ff;
}

.el-input__wrapper,
.el-textarea__inner,
.el-select__wrapper {
  border-radius: 10px !important;
}

.page-shell {
  max-width: 1240px;
  margin: 0 auto;
}

.muted {
  color: var(--app-muted);
}
</style>
