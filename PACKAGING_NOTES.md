# 打包记录

## 2026-04-14 继续验证

- 当前 `ch-2.ico` 文件头为标准 ICO：`00 00 01 00`
- 说明昨天记录中的“PNG 冒充 ICO”问题已经不再是当前仓库状态
- `build_with_signing.py` 已完成整链路验证：PyInstaller 打包成功、`signtool` 签名成功、签名校验通过
- 本次新增修复：`code_signer/sign_exe_file.py` 中原先使用了 `✓/✗/⚠`，在 Windows `gbk` 控制台下会触发 `UnicodeEncodeError`
- 现已改为 ASCII 状态标记（`[OK] / [FAIL] / [WARN] / [ERROR]`），避免签名完成后因打印输出异常中断
- 最新产物：
  - `dist/New_Customer_Operate.exe`
  - `signature_records/New_Customer_Operate.exe_signing_record.json`

## 2026-04-13 图标格式问题

- 文件: `ch-2.ico`
- 实际检查结果: 文件头是 PNG 签名 `89 50 4E 47 0D 0A 1A 0A`，不是标准 ICO 头 `00 00 01 00`
- 当前图片尺寸: `512 x 512`
- 直接后果: PyInstaller 在 Windows 下报错
  `ValueError: Received icon image ... is not in the correct format`

## 处理结论

- 当前仓库中的 `ch-2.ico` 实际上是“PNG 文件改了 .ico 扩展名”
- 没有安装 `Pillow` 时，PyInstaller 无法自动转换该文件
- 已在 `build_with_signing.py` 中增加前置检查
- 下次如果图标仍然不是有效 ICO，打包脚本会跳过 `--icon`，避免再次因为图标格式直接失败

## 下次修改建议

- 优先方案: 把 `ch-2.ico` 替换成真实的 Windows `.ico` 文件
- 备选方案: 保留 PNG，但显式保存为 `.png`，并在打包环境安装 `Pillow`
- 如果要兼容 Windows 图标显示，建议提供多尺寸 ICO，例如 `16/32/48/64/128/256`
