# 配置建议（config/）

本目录给出不依赖字节补丁即可落地的配置调整，优先执行（安全可逆）。

## 1. CUA 入口开启（computerUseComposerEntryHidden）

文件：`~/.zcode/v2/setting.json`

```json
{
  "computerUseComposerEntryHidden": false
}
```

将 `computerUseComposerEntryHidden` 改为 `false` 后重启 ZCode，
Composer 中 Computer Use 入口将不再隐藏（灰显缓解；功能可用性仍受 CUA 服务门禁影响，见 REPORT.md 第一节）。

> 说明：当前版本该字段默认 `true`（入口默认隐藏）。改为 `false` 为设置层修改，版本升级不覆盖用户设置。

## 2. 动态工作流：先确认是否有官方配置项

检查 `~/.zcode/v2/setting.json` 是否含：
- `dynamicWorkflow` / `dynamicWorkflowMode` / `workflowMode` 等键

若存在且可赋 `"onDemand"`，优先用配置方案（无需字节补丁）。
若不存在（当前 3.15.3 实测无），走 `scripts/enable-dynamic-workflow.ps1` 字节补丁方案。

## 3. 自动更新关闭（防补丁被覆盖）

- 设置界面关闭自动更新（或确认已关闭）
- 配套 `enable-dynamic-workflow.ps1` 会自动注册计划任务兜底

## 4. 内置插件离线同步

内网机首次部署时，从外网环境拷贝：
```
~/.zcode/cli/plugins/cache/          →  内网机相同路径
```
确保 computer-use 等内置插件在位（CUA 依赖 computer-use/0.6.4）。

## 5. 供应商配置（如需内网接自建/中转模型）

偏好设置 → 创建自定义供应商 → Chat Completions 格式 → Base URL 指向内网可达的
兼容 OpenAI 协议端点（如中转站 https://api.ai-pixel.online/v1），模型名小写（如 glm-5.3-flash）。
