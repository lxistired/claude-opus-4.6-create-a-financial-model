# 项目开发过程文档（谷歌财务模型）

## 1. 目标与范围
- 目标：基于**最近 4 年年报**（2022–2025 年 10-K），叠加网络公开信息，快速形成一个可解释、可复用的谷歌估值模型。
- 范围：先完成一版“公司整体 DCF”，暂不做复杂三表联动和分部联动。

## 2. 执行过程
1. 读取并定位年报关键章节（2022、2023、2024、2025 年 10-K）：
   - 合并利润表
   - 合并现金流量表
   - 资产负债表
   - 分部信息（Google Services / Cloud / Other Bets）
2. 提取 2022–2025 历史数据并计算核心比率（增长率、利润率、现金流率）。
3. 联网检索宏观锚点：
   - 10Y 美债收益率（用于 Rf）
   - 美联储 2% 通胀长期目标（用于终值增长逻辑）
4. 搭建 2026-2030 预测 + 终值 + 折现 + 净现金桥接。
5. 输出基准估值、敏感性矩阵、三情景结果。

## 3. 关键问题与处理
- 问题 1：联网结果对个股行情（beta/市值）抓取不稳定。
  - 处理：将 beta 与 ERP 作为可调参数，并在文档中显式披露假设来源与区间。
- 问题 2：2025 年 CapEx 大幅上行，若直接外推会压低短期 FCF。
  - 处理：采用“高位回落”路径（20% -> 16%）而非一次性归一化，保持谨慎。
- 问题 3：Other income 在 2025 年异常偏高，不宜作为经营常态。
  - 处理：使用 NOPAT（经营口径）作为 FCFF 起点，降低一次性项目干扰。

## 4. 结果摘要
- 基准每股价值：$121.4
- 敏感性区间（g=2.5%-3.5%，WACC=8.5%-10.0%）：$102.1 - $147.9
- 三情景：Bear/Base/Bull = $81.2 / $121.4 / $162.7

## 5. 复盘与教训
- 教训 1：估值模型先保证“透明可审计”，再追求复杂度；参数可解释性比参数精细度更重要。
- 教训 2：对 AI 周期下的高 CapEx 企业，必须将“资本开支路径”单独建模，而不是固定比例。
- 教训 3：网络信息波动较大时，应保留可替换参数并做敏感性分析，避免单点结论。

## 6. 下一步建议
- 升级为“分部驱动模型”：Services 与 Cloud 分别设定增长、利润率、CapEx 强度。
- 增加“回购路径”模块：将股本收缩纳入每股估值动态。
- 增加“监管罚款概率”情景：用于下修尾部风险估值。

## 7. 第一次迭代（投行版 Excel 交付）
- 新增 `Alphabet_IB_Model_2026.xlsx`（v1 简化版）。

## 8. 第二次迭代（分部驱动版 v2）
- 用户反馈 v1 太粗糙，未拆分业务线。
- 从头重建了 8-Tab 分部驱动模型 `Alphabet_IB_Model_v2.xlsx`：
  - Cover / Assumptions / Revenue_Build / P_and_L / Balance_Sheet / Cash_Flow / DCF / Sensitivity
- Revenue_Build 从 6 条业务线（Search、YouTube、Network、Subs、Cloud、Other Bets）自底向上拼总收入。
- Assumptions 支持 Base/Bull/Bear 三情景，CHOOSE 函数自动切换。
- 每条业务线独立设定 5 年增长率假设，可分情景调整。
- P&L 拆 TAC（按广告收入比例）+ Other COGS + R&D + S&M + G&A。
- DCF 从 EBIT 出发，扣税、加回 D&A、减 CapEx 和 ΔNWC，得 UFCF。
- Sensitivity 为 7×5 双变量（WACC × 终值 g）。
- 教训：Python `.format()` 与 Excel `{}` 数组冲突，需用 f-string 的 `{{}}` 转义。

## 9. 第三次迭代（分部 P&L 交叉验证版 v4）
- 用户反馈 v2 仍过于简单：缺少分部级成本拆解、CapEx 指引不对（应为 $180B）、假设缺乏历史数据锚点。
- 关键改进：
  1. **新增 Segment_PL tab**：Google Services / Cloud 各自拆出 Employee Comp + Other Costs（数据来自年报 Note 15）；Other Bets 用 OI margin 反推；Alphabet-level unallocated 单列
  2. **交叉验证行**：Segment_PL 中「Sum of Segment OI」与 Consolidated_PL 中「EBIT」比对，差异应=0
  3. **假设全面历史锚点**：Assumptions tab 每个参数都显示 2022A–2025A 历史值（灰色列）
  4. **CapEx 绝对值**：2026 使用管理层 Q4'25 earnings call 指引 $175B-$185B（中值 $180B），而非 % of revenue
  5. **术语脚注**：每个 Tab 底部添加 Glossary，解释所有缩写（TAC/COGS/SBC/UFCF/WACC/NOPAT/EV/D&A/OI&E/NWC/Ke/Kd）中英文对照
  6. **三情景 CHOOSE 增强**：Revenue Growth 拆为 Base/Bull/Bear 三个独立区块
- 产出文件：`Alphabet_IB_Model_v4.xlsx`（8 Tab）
- 生成脚本：`build_ib_model_v4.py`
- 技术问题：初始版本 `arow()` 函数处理变长参数时 `number_format` 被赋了 float 而非 string，导致 openpyxl 保存报错。修复：改为显式调用 arow() 传明确参数。
- 教训：
  - 投行模型必须体现年报实际可获取的最大粒度（如 Note 15 分部成本），而不是停留在合并口径
  - 每个假设都需要历史锚点，否则 reviewer 无法判断合理性
  - CapEx 当期有管理层指引时，应用绝对值输入而非公式驱动

## 10. 第四次迭代（三表联动版 v5 — 参照小米模型丰富度）
- **触发**：用户提供了一份小米财务模型 (`Xiaomi_claude_test.xlsx`) 作为参考，要求达到同等丰富度
- 小米模型特点：6 Tab (Key Summary / PnL / Drivers / BS / CF / Ratio Analysis)，每行有 Notes 列解释预测逻辑，底部有数据图例和来源

### 主要变更
1. **新增 Key_Summary tab**：执行摘要仪表板 — 核心KPI(Revenue/GP/EBIT/NI+Margins+YoY) + 分部收入结构+占比 + BS核心指标(Total Assets/Equity/Cash/ROE/ROA) + 投资亮点与风险
2. **新增 BS tab**：完整合并资产负债表
   - 流动资产: Cash(plug) + Marketable Securities + AR + Other
   - 非流动资产: Non-marketable Securities + Deferred Tax + PP&E(=Prior+CapEx-D&A) + Operating Lease + Goodwill + Other
   - 流动负债: AP + Accrued Comp + Accrued Expenses + Revenue Share + Deferred Revenue
   - 非流动负债: Long-term Debt + Tax NC + Operating Lease Liab + Other LT
   - 股东权益: Prior + NI - Buyback - Div + SBC(net)
   - **Balance Check 行**: Total Assets - (Total Liabilities + Equity) = 0
   - Cash 作为 plug 项确保恒等
3. **新增 Ratio_Analysis tab**：盈利能力(GPM/OPM/NPM/ROE/ROA) + 营运效率(AR Days/AP Days/CapEx Intensity/CapEx÷D&A) + 杠杆(D/A/Current Ratio/Net Cash) + 现金流质量(OCF÷NI/FCF÷Revenue) + 增长指标(Rev/NI/EPS YoY) + 投资亮点
4. **Notes 列 (column J)**：每个数据 tab 增加预测逻辑说明列
5. **数据图例 + 来源**：每个 tab 底部添加颜色编码说明(蓝字=原始/绿字=链接/蓝底=假设/黄底=关键) + 数据源引用
6. **增强 Cash_Flow**：WC 变动逐项拆分(AR/AP/Accrued Exp/Rev Share/Deferred Rev)，链接至 BS 变动
7. **增强 Assumptions**：新增 BS 相关假设(AR Days/AP Days/AccComp%/AccExp%/RevShare%/DefRev增速/OpLease增速/NonMktSec增速)，均附历史锚点
8. **修正**：2023 Cash & Marketable Securities 从 $95.7B 修正为 ~$110.9B（之前错误复制了2024值）
9. **BS 历史数据**：2024/2025 精确来自 10-K p.48；2023 根据公开数据合理估算（总资产≈$401B）

### 产出
- `Alphabet_IB_Model_v5.xlsx`（11 Tab 完整三表联动模型）
- `build_ib_model_v5.py`（~750行生成脚本）

### 教训
- 投行级模型的「丰富度」不仅是数据量，更在于：**每行有解释**(Notes列)、**颜色编码清晰**(数据图例)、**三表必须联动**(BS是关键纽带)、**综合比率分析**(Ratio tab 提供一目了然的健康度评估)
- Executive Summary (Key_Summary) 对决策者非常重要，投行 pitchbook 通常以此开头
- BS 的 plug 项(Cash) 是保证资产=负债+权益恒等的关键技巧
- 营运资金(WC)变动应从 BS 逐项推导而非笼统假设
