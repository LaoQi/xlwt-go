# xlwt-go

## 项目说明

纯 Go 实现的 .xls（Excel 97-2003，BIFF8 + OLE2 复合文档）**写入**库，从 Python [xlwt](https://github.com/python-excel/xlwt) 移植。包名 `xlwt`，模块 `github.com/LaoQi/xlwt-go`，`go.mod` 声明 `go 1.21`，零第三方依赖（仅标准库）。License: LGPL-2.1（与上游一致）。

面向使用方的公开 API（全部在包根，同一 package）：

- `NewWorkbook() *Workbook`
- `wb.AddSheet(name string) (*Worksheet, error)`：重名（大小写不敏感）、空名、超 31 字符、含 `\ / ? * [ ] :` 或前导 `'` 均报错
- `ws.Write(row, col int, label string) error`：仅字符串单元格；行列 0 基；键 `(row<<16)+col` 存 `Worksheet.Grid`，字符串进 `SST` 去重，XF 用 `DefaultCellXFStyle`(0x11)
- `ws.WriteWithStyle(r, c int, label string, style *XFStyle) error`：带样式写入，`nil` = 默认样式；`ws.Write` 就是它的无样式简写
- `wb.AddStyle(style *XFStyle) (int, error)`：注册样式并返回 XF 索引（`WriteWithStyle` 内部自动调用）
- 样式（`style.go`）：`XFStyle{NumFormatStr, Font, Alignment, Borders, Pattern, Protection}` + 各组件常量；`NewXFStyle()` 及各组件构造函数给出库默认值
- 尺寸（`dimensions.go`）：`SetColWidth`（单位=字符数，如 8.43）、`SetColWidthRaw`（1/256 字符宽）、`SetColDefaultWidth`（DEFCOLWIDTH）、`SetColHidden`、`SetRowHeight`（twips）/`SetRowHeightPoints`、`SetRowDefaultHeight`（DEFAULTROWHEIGHT）、`SetRowHidden`
- `wb.Save(w io.Writer) error`：`Workbook.GetBiffData()` 产出 BIFF 流 → `XlsDoc.Save` 包成 OLE2 写出。任意 `io.Writer`（文件、`bytes.Buffer`、HTTP body）均可
- 校验（`errors.go`）：上限常量 `MaxRow`=65535、`MaxCol`=255、`MaxStringLength`=32767（按**字符**计，非字节）、`MaxSheetNameLength`=31；哨兵错误 `ErrRowOutOfRange`/`ErrColOutOfRange`/`ErrStringTooLong`/`ErrInvalidSheetName`/`ErrDuplicateSheetName`，用 `%w` 包装，配 `errors.Is` 使用。**出错时不写入任何状态**，工作簿里永远不会出现 .xls 无法表达的数据

文件职责：

- `doc.go`：包级文档（pkg.go.dev 首页说明）
- `example_test.go`：外部包 `xlwt_test` 的可运行示例（godoc 展示）
- `workbook.go` / `worksheet.go`：模型与 BIFF 记录组装顺序（改记录顺序看这里）
- `sst.go`：SharedStringTable，SST + CONTINUE(0x3C) 分块（`MaxSSTLength`=0x2020）；超长字符串按上游语义**拆到后续 CONTINUE**，续块首字节重复选项标志（`MaxSSTCellLength` 已弃用）
- `style.go`：样式数据类型（`Font`/`Alignment`/`Borders`/`Pattern`/`Protection`/`XFStyle`）、全部相关常量、内置数字格式表（`stdNumFormatStrings`，前 23 项 = 索引 0..22，后 13 项 = 37..49）
- `style_collection.go`：索引分配 + 去重 + 样式段序列化。**索引布局刻意复刻上游**：字体 0,1,2,3,5,6,7（4 号在所有 BIFF 版本都跳过）、style XF 0x00..0x0F（指向 6 号字体）、cell XF 自 0x10 起 → 默认样式落在 0x11。上游用对象身份去重（每个等价样式的副本一个新 XF），本库**按值去重**（更省，等价样式/字体合并）
- `record_style.go`：`fontRecord`/`numberFormatRecord`/`xfRecord`/`styleRecord` 参数化实现，位域打包严格对齐上游 `XFRecord`（含"无边框线则颜色写 0"这一细节）
- **硬约束**：默认路径（不传样式）的输出字节必须与加样式功能之前**完全一致**（`sha256 = 50B78E44…706DE`，由 `TestStyle_DefaultOutputUnchanged` 与字节级对照保障）
- `biff.go`：`BiffRecord`（记录头；数据 >0x2020 自动拆 CONTINUE）、`SingleHRecord`
- `record_workbook.go` / `record_worksheet.go` / `record_style.go`：各 BIFF 记录序列化，记录号硬编码（0x0809 BOF、0x00FD LABELSST、0x0208 Row、0x00E0 XF…）
- `compound_doc.go`：OLE2 容器（Header、MSAT/SAT、Directory），512 字节扇区；内部中文注释是乱码，勿依赖
- `types.go`：`SP_L/SP_l/SP_H/SP_h/SP_I/SP_B/SP_d` 对应 `struct.pack` 的定长整型别名；`util.go`：UTF-16LE / ASCII 打包与填充

已知限制与待办（提需求/改代码前先看）：

- 只支持字符串单元格；无数字/日期/布尔/公式；`BlankRecord` 已实现但未使用
- 单个字符串上限 32767 字符：现由 `Write` 校验并返回 `ErrStringTooLong`（此前会静默产出损坏文件）
- 未实现/被忽略：Palette、Password(0x13)、Backup、Window1、Country/Links，见源码 `@fixme` / `@todo`
- `record_workbook.go` `BoundSheetRecord` 已修（2026-09）：此前 sheet 名一律按 8-bit 压缩格式打包（`ASCIIStringPack`），导致非 latin-1 表名（中文/日文/西里尔等）乱码。现改用 `U16StringPack1Byte`，按上游 `upack1` 语义在 latin-1 与 UTF-16LE 间自动选择
- `workbook.go` `BoundsSheetsRec` 已修（2026-09）：此前所有 BOUNDSHEET 写入**同一个**流偏移，导致多 sheet 工作簿里第三方读取器（xlrd 等）读每个 sheet 都返回**第一个 sheet 的内容**。现按上游 `start += sheet_biff_len` 逐 sheet 递增
- `compound_doc.go` 已修两处 OLE2 容器缺陷（2026-09）：
  - `BuildSat` 的 MSAT 二级扇区循环此前**漏掉上游的 `i += 1`**（该自增只存在于 else 分支），导致 BIFF 流 ≥ 约 6.76MB（`SAT_sect_count` > 109）时 panic `index out of range`。已按 Python xlwt 补回，并有回归测试 `TestXlsDoc_SaveLargeStream`（8MiB 流，已验证无修复时必失败）
  - `book_stream_sect = append(dir_stream_sect, sect)` 目标写错 → 改为 `append(xls.book_stream_sect, ...)`；该切片不参与输出，修复前后输出字节哈希一致
  - 验证口径：同一 stream 喂 Go/Python 的 `XlsDoc`，512B~33MB（含阈值与二级 MSAT 区间）输出**逐字节一致**
- 复用同一 `XlsDoc`/`Workbook` 实例重复 `Save` 会累积扇区状态并产出损坏文件（上游 Python 亦然，属同源缺陷，未改）；每次 `Save` 请用新实例
- `Write`/`AddSheet` 已加 error 返回（2026-09，属破坏性 API 变更，但项目此前无 tag）；`Save` 仍不校验工作簿整体大小；库内不打印日志，仅通过返回值暴露错误
- 样式已支持字体/数字格式/对齐/边框/填充/保护，但**未实现**上游 `easyxf`/`easyfont` 字符串 DSL（阶段 2）与 `add_palette_colour` 自定义调色板（阶段 3）
- 列宽/行高已支持（2026-09，对齐上游）：**只有显式设置过的列/行才写出 COLINFO/ROW 记录**，未设置时输出字节与加该功能前完全一致（同样受硬约束与 `TestDimensions_UnaffectedByDefault` 保护）
- 上游口径细节（已实证核对，勿凭直觉改）：ROW 记录 bit15「使用默认行高」上游**始终为 0**（即使高度就是默认值）；bit 27-16 为默认 XF 索引 `0x0F`（ROW 与 COLINFO 都用 0x0F）；bit8 恒为 1。GUTS 的行层级在无分行时也报 `1`，列层级取已用 COLINFO 的 `max(level)+1`
- 写单元格**不会**隐式创建 COLINFO（上游 `ws.col(i).width=` 才创建）；只设行属性而无单元格的行也会输出 ROW 记录
- 尚未支持：单元格高亮、合并单元格；行列的分级（outline level）记录已具备但未暴露 API
- 零值语义有 2 处例外（文件格式把 0 用作有意义的非默认值）：`Alignment.Vert`（0=顶端）与 `Protection.CellUnlocked`（0=锁定，恰为默认）。其余字段零值回落到文档所述默认。**需要库默认值时请用 `NewXFStyle()` 或组件构造函数**
- 导出面偏大：BIFF/OLE2 记录层（`LabelSSTRecord`、`Window1Record`、`XlsDoc`…）也暴露在 godoc 中，如需收敛可迁到 `internal/`（较大重构，未做）

## 构建与测试

```powershell
gofmt -l .          # 应无输出
go build ./...
go vet ./...        # 当前无告警
go test ./...       # 全部测试，约 1s
go test -race ./... # 需 cgo/gcc；本机 Windows 无 gcc 会 build failed，CI 中仅 linux/mac 跑 -race
go test -run TestWorksheet_Write -v .
```

- 测试写入 `t.TempDir()`（不再污染包目录）。除结构断言（OLE2 签名、扇区对齐）外，测试内含一个小型 BIFF 解析器（`parseBiffRecords`/`sstPayloads`/`parseSSTStrings`/`sheetCells`），可**内容级**校验：SST 表布局与 CONTINUE 拆分、BOUNDSHEET 偏移指向各自 BOF、各 sheet 单元格归属正确
- 回归测试与被修复的 bug 一一对应，可放心回退验证：`TestWorkbook_BoundSheet*`（偏移不递增）、`TestSharedStringTable_LongStringsUsedToBeDropped`（长串被丢弃）、`TestXlsDoc_SaveLargeStream`（MSAT 二级扇区 panic）、`TestWorkbook_NonLatinSheetNameUsedToBeMangled`（表名乱码）、`TestWorksheet_WriteValidation` / `TestWorkbook_AddSheetValidation`（参数校验）
- 外部读取器交叉验证（Python xlrd 2.0.1 实测通过）：多 sheet 内容各归其位、5000 字符/emoji 长串、中文与其他 Unicode、3000 条 SST 跨 CONTINUE、8000 格 33MB 走 MSAT 二级路径的大文件
- 行尾统一 LF（`.gitattributes` 强制 `eol=lf`），`gofmt -l .` 不再误报；`.editorconfig` 约定 Go 用 tab
- CI：`.github/workflows/ci.yml`，linux/windows/macos × gofmt+vet+build+test，另加 govulncheck
- 端口对照原则：改动尽量对齐 Python xlwt 同名方法的字节布局，注释里的 `struct.pack` 就是对照说明
- 发布：`go get` 依赖 tag，例如 `git tag v0.1.0 && git push --tags`（尚未打 tag）