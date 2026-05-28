# 公开发布合规性检查报告

检查日期：2026-05-28  
对象：`技术自由主义` 论文项目的 GitHub 公开展示包

## 结论

原始项目目录不建议原样公开。公开包可以发布，但应以“透明展示 AI 写作过程与修订机制”的性质发布，而不是宣称原始目录中的每个文件都已经适合公开。

主要理由：

- 原始目录含第三方 PDF/EPUB 附件，不应直接上传公开仓库。
- 原始目录含旧版 Word/PPT、系统残留、缓存和导出模板文件，不适合作为公开展示主体。
- 原始 `chapters/` 中仍保留早期核查清单标红的若干 AI 幻觉引用；公开包副本已做最小修正。
- 过程材料中存在导师、评审和协作语境信息；本公开包保留与论文透明度直接相关的记录，但未包含原始会议转录。

## 已做的最小修正

公开包副本相对原始 `chapters/` 做了以下修正：

1. 删除第 2 章中 Heller 2010 的虚构期刊文章引用和引介句。
2. 删除第 2 章中 Desmond et al. 2015 的两处虚构来源引用，并从参考文献中删除对应条目。
3. 删除第 4 章中“胡翌霖 2020a + 芒福德 + 中国传统”的脚注。
4. 将第 2 章 “What Medium Can Mean” 引用从 `(Rancière, 2010: 36)` 修正为 `(Rancière, 2011: 36)`，参考文献补入 *Parrhesia* 11:35-43。
5. 将 Rivas 条目修正为 `Rivas, 2025, Educational Philosophy and Theory, 57(5):435-449`，并同步正文引用为 `(Rivas, 2025)`。
6. 删除第 2 章中与朗西埃学术历程分期不匹配的 `(蒋洪生, 2012)` 并保留吕峰 2024 作为中文研究参照。

外部核验依据：

- Rancière, “What Medium Can Mean,” *Parrhesia* 11, 2011: https://www.parrhesiajournal.org/parrhesia11/parrhesia11_ranciere.pdf
- Taylor & Francis 卷期页显示 Rivas 文章为 *Educational Philosophy and Theory* 57(5):435-449: https://www.tandfonline.com/toc/rept20/57/5
- PhilPapers 条目同样记录 Rivas DOI 与 57(5):435-449: https://philpapers.org/rec/RIVBSA

## 已排除内容

- 第三方 PDF/EPUB 文献附件。
- Word、PPT、旧版导出和模板 `reference.docx`。
- `.DS_Store`、`__pycache__`、锁文件、FUSE 残留。
- 原始会议转录。
- 私有云盘目录与附件下载包。

## 仍需作者确认的边界

- 是否希望公开过程日志中出现的导师、评审、公司与协作者名称。
- 是否希望将论文正文改为匿名版或保留署名版。
- 是否需要在学校/学院层面确认“学位论文授权说明”与 GitHub 自主公开之间的关系。
- 是否需要将 Word/PDF 正式版重新生成并做元数据清理后再加入 release，而不是放入仓库主分支。

## 合规建议

当前公开包适合先作为 GitHub 仓库公开展示。如果后续要发布正式 PDF 或 Word 版本，建议另行执行：

- 从已修正 Markdown 重新导出；
- 清理 Office 元数据；
- 检查目录、页码、脚注和参考文献；
- 只作为 GitHub Release 附件发布，不与第三方文献混放。

