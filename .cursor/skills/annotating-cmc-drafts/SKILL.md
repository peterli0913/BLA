---
name: annotating-cmc-drafts
description: Use when reviewing a PPQ protocol, ESA, gap assessment, MOR, or parameter-range workbook for 批注, inconsistencies, mistranslation, or errors, or when checking that PVP, Excel parameter ranges, and MOR process parameters match.
---

# 批注 CMC 草稿

对照锁定终版或同系列已定稿。只报会改变执行、设备结论或文件引用的差异。

## 输出

不改源 Word，不在 Word 里插入批注。新建 Excel，四列：序号、位置、内容（可在 Word 里原样搜索的短句）、批注内容（错在哪里、改成什么）。一行一件事。同一类重复错误并成一行，搜索句给第一处，批注里写还会在哪些步骤重复。写完用脚本校验每个搜索句都能在接受修订后的文本里命中。

## 先看这次改了什么

有上一版时，先对接受修订后的文本做差。已经改掉的项不要再报。改了一处但同句或另一节还留着旧字符串的，要再搜一遍。相邻两版还要查回退：上一版改对、这一版又变回去的（如 αMePhe → aMePhe、10.4 V → 10.39 V），单独报。

客户批注按邮箱认，不按姓名看起来像哪国人。`@lilly.com` 和 `@network.lilly.com` 是客户。Yiwen Zhang（yiwen.zhang@lilly.com）是礼来客户。`done=0` 且要求补 IPC、改地址或改数字的，正文没答之前继续报。客户后来澄清只是问小数位的，不要把前一条猜测写成必须改的数。

## 三方参数核对（PVP / 参数表 / MOR）

参数表是客户 PAR 和 PFD 转出来的，定了 NOR；PVP 引用参数表；MOR 是执行文件。

1. 先读 MOR 的原料信息表（按操作号列出物料、V、kg/kg、标准投料 kg、称量范围），这是对投料最快的入口。再读正文：MOR 正文和原料信息表可能一个改了一个没改。
2. 参数表的 kg/kg × 树脂投料（L2 Step 1 为 48 kg）换成 kg，再对 MOR。参数表 NOR 多为 ±5%（个别 ±4%），MOR 称量范围多为 ±3%。
3. 抽出 MOR 所有 ★ 句和“within the Normal Operating Range (NOR)”句里的数，与 PVP 表 7 逐条比。PVP 表头写“NOR (MOR)”时，两边数字应相同。
4. 对 CPP 逐个问一句：MOR 在这一步有没有控制和记录点。L2 的“溶解夹套温度”在 PVP 是 CPP，MOR 在溶解阶段没有控温。
5. 偶联温度、偶联总时间上限、脱保护时间、洗涤体积和次数，按氨基酸逐个核，不要按“标准 AA”一句带过。
6. 参数表行名与 MOR 操作号对不准的，先给出判断和把握，单独一张表，最后集中问用户。

算产量前先看 MOR 的理论产量公式。L2 Step 1 的产量范围随树脂取代度变化，不是投料 × 收率。

## 取舍

报：单位错；同一文件里两套数；三份文件之间不一致的 CPP、投料体积、时间和温度；质量标准号张冠李戴；客户要求列入的 IPC 仍缺；会执行错的判据；文件号前后不一；Target 栏写成 PAR 下限。

不报：终版反复留着且不影响执行的语病，除非用户要求“从头捋”；只差 25 与 25.0；客户已标 done 的提问。

批注写改法。
