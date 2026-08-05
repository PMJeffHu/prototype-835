const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, WidthType, BorderStyle, ShadingType,
  convertInchesToTwip, LevelFormat, TabStopPosition, TabStopType
} = require('docx');
const fs = require('fs');

// ===== Helper functions =====
function heading(text, level) {
  return new Paragraph({
    heading: level,
    spacing: { before: 360, after: 200 },
    children: [new TextRun({ text, bold: true, font: 'Microsoft YaHei' })]
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, font: 'Microsoft YaHei', size: 21, ...opts })]
  });
}

function codeBlock(lines) {
  const children = [];
  lines.forEach((line, i) => {
    children.push(
      new TextRun({ text: line, font: 'Consolas', size: 18 })
    );
    if (i < lines.length - 1) children.push(new TextRun({ break: 1 }));
  });
  return new Paragraph({
    spacing: { after: 120, before: 80 },
    shading: { type: ShadingType.SOLID, color: 'F5F5F5', fill: 'F5F5F5' },
    indent: { left: convertInchesToTwip(0.3) },
    children
  });
}

function boldPara(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, font: 'Microsoft YaHei', size: 21, bold: true, ...opts })]
  });
}

function createTable(headers, rows) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map(h =>
      new TableCell({
        shading: { type: ShadingType.SOLID, color: '4472C4', fill: '4472C4' },
        width: { size: 4500, type: WidthType.DXA },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: h, bold: true, color: 'FFFFFF', font: 'Microsoft YaHei', size: 20 })]
        })]
      })
    )
  });

  const dataRows = rows.map(row =>
    new TableRow({
      children: row.map((cell, ci) =>
        new TableCell({
          width: { size: 4500, type: WidthType.DXA },
          shading: ci === 0 ? { type: ShadingType.SOLID, color: 'E8EEF7', fill: 'E8EEF7' } : undefined,
          children: [new Paragraph({
            children: [new TextRun({ text: cell, font: 'Microsoft YaHei', size: 20 })]
          })]
        })
      )
    })
  );

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [headerRow, ...dataRows]
  });
}

// ===== Build Document =====
const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: 'Microsoft YaHei', size: 21 }
      }
    }
  },
  sections: [{
    properties: {
      page: {
        margin: { top: convertInchesToTwip(0.8), bottom: convertInchesToTwip(0.8), left: convertInchesToTwip(1), right: convertInchesToTwip(1) }
      }
    },
    children: [

      // ========== TITLE ==========
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
        children: [
          new TextRun({ text: '发票与账单自动核销算法设计文档', bold: true, size: 32, font: 'Microsoft YaHei' })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 600 },
        children: [
          new TextRun({ text: '2026-07-24', size: 21, font: 'Microsoft YaHei', color: '666666' })
        ]
      }),

      // ========== 目录标题 ==========
      heading('目  录', HeadingLevel.HEADING_1),
      para('一、数据结构与枚举定义'),
      para('二、汇率获取接口'),
      para('三、主算法伪代码'),
      para('四、辅助函数'),
      para('五、算法核心流程图解'),
      para('六、边界情况处理'),

      // ========== 一、数据结构 ==========
      heading('一、数据结构与枚举定义', HeadingLevel.HEADING_1),
      heading('1.1 枚举定义', HeadingLevel.HEADING_2),

      heading('开票币种规则', HeadingLevel.HEADING_3),
      codeBlock([
        'enum BillingCurrencyRule {',
        '    BILL_CURRENCY,          // 账单币种',
        '    SPECIFIED_CURRENCY      // 指定币种',
        '}',
      ]),

      heading('发票类型', HeadingLevel.HEADING_3),
      codeBlock([
        'enum InvoiceType {',
        '    VAT_INVOICE,            // 增值税发票',
        '    PROFORMA_INVOICE        // 形式发票',
        '}',
      ]),

      heading('汇率规则', HeadingLevel.HEADING_3),
      codeBlock([
        'enum ExchangeRateRule {',
        '    BOOK_RATE,              // 记账汇率',
        '    TRANSFER_RATE,          // 转汇汇率',
        '    SPECIFIED_RATE          // 指定汇率（需移除，不参与自动核销）',
        '}',
      ]),

      heading('核销状态', HeadingLevel.HEADING_3),
      codeBlock([
        'enum WriteOffStatus {',
        '    PENDING,                // 待核销',
        '    PARTIAL,                // 部分核销',
        '    COMPLETED               // 已核销',
        '}',
      ]),

      heading('自动参与标识', HeadingLevel.HEADING_3),
      codeBlock([
        'enum AutoParticipate {',
        '    PARTICIPATE,            // 参与',
        '    NOT_PARTICIPATE         // 不参与',
        '}',
      ]),

      heading('1.2 实体定义', HeadingLevel.HEADING_2),

      heading('客户开票流程配置', HeadingLevel.HEADING_3),
      codeBlock([
        'struct CustomerInvoiceConfig {',
        '    customerAccountCode:   string;              // 客户账户编号',
        '    autoAssociateBill:     boolean;             // 发票自动关联账单',
        '    billingCurrencyRule:   BillingCurrencyRule; // 开票币种规则',
        '    specifiedCurrency:     string | null;       // 指定币种',
        '    invoiceType:           InvoiceType;         // 发票类型',
        '    exchangeRateRule:      ExchangeRateRule;    // 汇率规则',
        '    specifiedRate:         decimal | null;      // 指定汇率',
        '}',
      ]),

      heading('发票台账记录', HeadingLevel.HEADING_3),
      codeBlock([
        'struct Invoice {',
        '    invoiceNo:            string;         // 发票号',
        '    invoiceDate:          date;           // 发票日期',
        '    invoiceCurrency:      string;         // 发票币种',
        '    invoiceAmount:        decimal;        // 发票金额',
        '    writeOffableAmount:   decimal;        // 可核销金额（可正可负）',
        '    customerAccountCode:  string;         // 客户账户编号',
        '    writeOffStatus:       WriteOffStatus; // 核销状态',
        '    autoParticipate:      AutoParticipate;// 自动参与核销标识',
        '    invoiceType:          InvoiceType;    // 发票类型',
        '}',
      ]),

      heading('对账管理-账单记录', HeadingLevel.HEADING_3),
      codeBlock([
        'struct Bill {',
        '    billNo:               string;    // 账单号',
        '    billPeriod:           string;    // 账单周期',
        '    billAccountingDate:   date;      // 账单账务日期',
        '    billCurrency:         string;    // 账单币种',
        '    billableAmount:       decimal;   // 可开票金额（账单币种）',
        '    writeOffableAmount:   decimal;   // 可核销金额（不为0）',
        '    customerAccountCode:  string;    // 客户账户编号',
        '    exchangeRate:         decimal;   // 发票汇率（账单币种→开票币种）',
        '}',
      ]),

      heading('待核销账单（步骤4计算后的中间结构）', HeadingLevel.HEADING_3),
      codeBlock([
        'struct PendingBill {',
        '    billNo:                    string;   // 账单号',
        '    billPeriod:                string;   // 账单周期',
        '    billAccountingDate:        date;     // 账单账务日期',
        '    billCurrency:              string;   // 账单币种',
        '    billingCurrency:           string;   // 开票币种',
        '    exchangeRate:              decimal;  // 汇率（账单币种→开票币种）',
        '    billableAmountInBillingCurr: decimal; // 开票币种下的可开票金额',
        '    remainingBillableAmount:   decimal;  // 剩余可分配金额',
        '    customerAccountCode:       string;   // 客户账户编号',
        '}',
      ]),

      heading('核销任务', HeadingLevel.HEADING_3),
      codeBlock([
        'struct WriteOffTask {',
        '    taskId:               string;            // 任务ID',
        '    invoiceNo:            string;            // 发票号',
        '    writeOffAmount:       decimal;           // 本次核销金额（开票币种）',
        '    writeOffCurrency:     string;            // 核销币种（开票币种）',
        '    writeOffDate:         date;              // 核销日期',
        '    billAllocations:      List<BillAllocation>; // 账单分配明细',
        '}',
      ]),

      heading('账单分配明细', HeadingLevel.HEADING_3),
      codeBlock([
        'struct BillAllocation {',
        '    billNo:                    string;   // 账单号',
        '    billCurrency:              string;   // 账单币种',
        '    allocatedAmountInBillCurr:  decimal;  // 分配金额（账单币种）',
        '    allocatedAmountInBillingCurr: decimal; // 分配金额（开票币种）',
        '    exchangeRate:              decimal;  // 使用汇率',
        '}',
      ]),

      // ========== 二、汇率获取接口 ==========
      heading('二、汇率获取接口（按客户规则获取）', HeadingLevel.HEADING_1),
      heading('2.1 获取汇率函数', HeadingLevel.HEADING_2),
      codeBlock([
        'function getExchangeRate(',
        '    customerConfig: CustomerInvoiceConfig,',
        '    billCurrency:    string,',
        '    billingCurrency: string,',
        '    accountingDate:  date',
        '): decimal {',
        '    if billCurrency == billingCurrency:',
        '        return 1.0',
        '',
        '    switch customerConfig.exchangeRateRule:',
        '        case BOOK_RATE:',
        '            // 从汇率表按账务日期取记账汇率',
        '            return fetchBookRate(billCurrency, billingCurrency, accountingDate)',
        '        case TRANSFER_RATE:',
        '            // 从汇率表按账务日期取转汇汇率',
        '            return fetchTransferRate(billCurrency, billingCurrency, accountingDate)',
        '        // SPECIFIED_RATE 的客户已在步骤1移除',
        '}',
      ]),

      heading('2.2 币种转换函数', HeadingLevel.HEADING_2),
      codeBlock([
        'function convertCurrency(',
        '    amount:           decimal,',
        '    fromCurrency:     string,',
        '    toCurrency:       string,',
        '    rate:             decimal',
        '): decimal {',
        '    if fromCurrency == toCurrency:',
        '        return amount',
        '    return amount * rate',
        '}',
      ]),

      // ========== 三、主算法 ==========
      heading('三、主算法伪代码', HeadingLevel.HEADING_1),

      heading('主入口函数', HeadingLevel.HEADING_2),
      codeBlock([
        'function autoWriteOffInvoicesWithBills(): List<WriteOffTask> {',
      ]),

      // Step 1
      heading('步骤1：构造自动核销客户列表', HeadingLevel.HEADING_3),
      codeBlock([
        '    customerList = []  // List<CustomerInvoiceConfig>',
        '',
        '    // 1a. 查询所有"发票自动关联账单=是"的客户账户编号',
        '    allConfigs = queryCustomerInvoiceConfig(',
        '        where: { autoAssociateBill: true }',
        '    )',
        '',
        '    // 1b. 过滤：移除汇率规则为"指定汇率"的客户',
        '    for each config in allConfigs:',
        '        if config.exchangeRateRule == SPECIFIED_RATE:',
        '            continue  // 指定汇率的客户不参与自动核销，跳过',
        '        customerList.append(config)',
        '',
        '    if customerList.isEmpty():',
        '        log("无符合条件的自动核销客户，流程结束")',
        '        return []',
        '',
        '    // 提取客户账户编号集合',
        '    autoCustomerCodes = customerList',
        '        .map(c => c.customerAccountCode).toSet()',
      ]),

      // Step 2
      heading('步骤2：构造待核销发票列表', HeadingLevel.HEADING_3),
      codeBlock([
        '    rawInvoices = queryInvoiceLedger(',
        '        where: {',
        '            customerAccountCode IN autoCustomerCodes,',
        '            writeOffStatus IN [PENDING, PARTIAL],',
        '            autoParticipate == PARTICIPATE,',
        '            writeOffableAmount != 0',
        '        }',
        '    )',
        '',
        '    // 按发票日期从早到晚排序',
        '    pendingInvoices = rawInvoices.sortBy(',
        '        invoice => invoice.invoiceDate, ASC)',
      ]),

      // Step 3
      heading('步骤3：构造待核销账单列表（原始）', HeadingLevel.HEADING_3),
      codeBlock([
        '    rawBills = queryReconciliationBill(',
        '        where: {',
        '            customerAccountCode IN autoCustomerCodes,',
        '            writeOffableAmount != 0',
        '        }',
        '    )',
        '',
        '    // 按账单账务日期从早到晚排序',
        '    rawBills = rawBills.sortBy(',
        '        bill => bill.billAccountingDate, ASC)',
      ]),

      // Step 4
      heading('步骤4：计算开票币种下的可开票金额', HeadingLevel.HEADING_3),
      codeBlock([
        '    pendingBills = []  // List<PendingBill>',
        '',
        '    for each bill in rawBills:',
        '        // 4a. 找到该账单对应的客户配置',
        '        customerConfig = customerList.find(',
        '            c => c.customerAccountCode == bill.customerAccountCode)',
        '',
        '        // 4b. 确定开票币种',
        '        switch customerConfig.billingCurrencyRule:',
        '            case BILL_CURRENCY:',
        '                billingCurrency = bill.billCurrency',
        '            case SPECIFIED_CURRENCY:',
        '                billingCurrency = customerConfig.specifiedCurrency',
        '',
        '        // 4c. 获取汇率（账单币种 → 开票币种）',
        '        rate = getExchangeRate(customerConfig,',
        '            bill.billCurrency, billingCurrency, bill.billAccountingDate)',
        '',
        '        // 4d. 计算开票币种下的可开票金额',
        '        billableAmountInBillingCurr = convertCurrency(',
        '            bill.writeOffableAmount,',
        '            bill.billCurrency, billingCurrency, rate)',
        '',
        '        // 构造待核销账单对象',
        '        pendingBill = PendingBill{',
        '            billNo:                    bill.billNo,',
        '            billAccountingDate:        bill.billAccountingDate,',
        '            billCurrency:              bill.billCurrency,',
        '            billingCurrency:           billingCurrency,',
        '            exchangeRate:              rate,',
        '            billableAmountInBillingCurr: billableAmountInBillingCurr,',
        '            remainingBillableAmount:   billableAmountInBillingCurr,',
        '            customerAccountCode:       bill.customerAccountCode,',
        '            ...',
        '        }',
        '        pendingBills.append(pendingBill)',
      ]),

      // Step 5-6
      heading('步骤5-6：循环匹配发票与账单，创建核销任务', HeadingLevel.HEADING_3),
      codeBlock([
        '    writeOffTasks = []     // List<WriteOffTask>',
        '    billIndex = 0          // 账单列表FIFO消费指针',
        '',
        '    for each invoice in pendingInvoices:',
        '',
        '        remainingInvoiceAmount = invoice.writeOffableAmount',
        '',
        '        while remainingInvoiceAmount != 0',
        '              AND billIndex < pendingBills.length:',
      ]),

      // 5a
      boldPara('步骤5a：发票可核销金额为负数时'),
      codeBlock([
        '            if remainingInvoiceAmount < 0:',
        '',
        '                // 跳过正数账单',
        '                if pendingBills[billIndex].remainingBillableAmount > 0:',
        '                    billIndex = billIndex + 1',
        '                    continue',
        '',
        '                // 累积负数账单直到汇总为负',
        '                accumulatedBills = []',
        '                accumulatedAmount = 0',
        '',
        '                while billIndex < pendingBills.length',
        '                      AND accumulatedAmount >= 0:',
        '                    bill = pendingBills[billIndex]',
        '',
        '                    if bill.remainingBillableAmount > 0:',
        '                        billIndex++  // 跳过正数账单',
        '                        continue',
        '',
        '                    accumulatedBills.append(bill)',
        '                    accumulatedAmount += bill.remainingBillableAmount',
        '                    billIndex++',
        '',
        '                if accumulatedAmount >= 0:',
        '                    break  // 无可匹配的负数账单',
        '',
        '                // 取绝对值较小者作为本次核销金额',
        '                absInvoice  = abs(remainingInvoiceAmount)',
        '                absAccum    = abs(accumulatedAmount)',
        '',
        '                if absAccum <= absInvoice:',
        '                    writeOffAmount = accumulatedAmount',
        '                else:',
        '                    writeOffAmount = remainingInvoiceAmount',
      ]),

      // 5b
      boldPara('步骤5b：发票可核销金额为正数时'),
      codeBlock([
        '            else:  // remainingInvoiceAmount > 0',
        '',
        '                // 跳过非正数账单',
        '                bill = pendingBills[billIndex]',
        '                if bill.remainingBillableAmount <= 0:',
        '                    billIndex++',
        '                    continue',
        '',
        '                // 累积正数账单',
        '                accumulatedBills = []',
        '                accumulatedAmount = 0',
        '',
        '                while billIndex < pendingBills.length:',
        '                    bill = pendingBills[billIndex]',
        '',
        '                    if bill.remainingBillableAmount <= 0:',
        '                        break  // 遇到非正数账单停止',
        '',
        '                    accumulatedBills.append(bill)',
        '                    accumulatedAmount += bill.remainingBillableAmount',
        '                    billIndex++',
        '',
        '                    if accumulatedAmount >= remainingInvoiceAmount:',
        '                        break  // 已累积足够金额',
        '',
        '                if accumulatedAmount <= 0:',
        '                    break  // 无可匹配的正数账单',
        '',
        '                // 取较小值',
        '                if accumulatedAmount <= remainingInvoiceAmount:',
        '                    writeOffAmount = accumulatedAmount',
        '                else:',
        '                    writeOffAmount = remainingInvoiceAmount',
      ]),

      // 创建任务
      boldPara('创建核销任务并分配金额'),
      codeBlock([
        '            // 创建核销任务',
        '            task = createWriteOffTask(',
        '                invoice, accumulatedBills,',
        '                writeOffAmount, billingCurrency)',
        '            writeOffTasks.append(task)',
        '',
        '            // 更新发票剩余可核销金额',
        '            remainingInvoiceAmount -= writeOffAmount',
        '',
        '            // 回退未完全消耗的账单指针',
        '            rollbackUnconsumedBills(',
        '                pendingBills, accumulatedBills,',
        '                writeOffAmount, billIndex)',
        '',
        '        end while  // 发票核销循环结束',
        '    end for      // 所有发票处理完毕',
      ]),

      // 持久化
      boldPara('持久化核销任务'),
      codeBlock([
        '    for each task in writeOffTasks:',
        '        persistWriteOffTask(task)',
        '        updateInvoiceAfterWriteOff(',
        '            task.invoiceNo, task.writeOffAmount)',
        '        for each allocation in task.billAllocations:',
        '            updateBillAfterWriteOff(',
        '                allocation.billNo,',
        '                allocation.allocatedAmountInBillCurr)',
        '',
        '    return writeOffTasks',
        '}',
      ]),

      // ========== 四、辅助函数 ==========
      heading('四、辅助函数', HeadingLevel.HEADING_1),

      heading('4.1 创建核销任务', HeadingLevel.HEADING_2),
      codeBlock([
        'function createWriteOffTask(',
        '    invoice:          Invoice,',
        '    accumulatedBills: List<PendingBill>,',
        '    writeOffAmount:   decimal,     // 开票币种下',
        '    billingCurrency:  string',
        '): WriteOffTask {',
        '',
        '    billAllocations = []',
        '    remainingToAllocate = abs(writeOffAmount)',
        '    sign = writeOffAmount > 0 ? 1 : -1',
        '',
        '    // FIFO按比例分配到各账单',
        '    for each bill in accumulatedBills:',
        '        if remainingToAllocate <= 0:',
        '            break',
        '',
        '        billAbsAmount = abs(bill.remainingBillableAmount)',
        '        totalAbsAmount = sumOfAbs(accumulatedBills)',
        '',
        '        if totalAbsAmount == 0: continue',
        '',
        '        // 按比例分配（开票币种）',
        '        allocatedInBillingCurr =',
        '            sign * billAbsAmount * remainingToAllocate / totalAbsAmount',
        '',
        '        // 确保不超账单剩余金额',
        '        allocatedInBillingCurr = minByAbs(',
        '            allocatedInBillingCurr, bill.remainingBillableAmount)',
        '',
        '        // 转换为账单币种',
        '        if bill.exchangeRate != 0:',
        '            allocatedInBillCurr = allocatedInBillingCurr / bill.exchangeRate',
        '        else:',
        '            allocatedInBillCurr = allocatedInBillingCurr',
        '',
        '        allocation = BillAllocation{',
        '            billNo:                    bill.billNo,',
        '            billCurrency:              bill.billCurrency,',
        '            allocatedAmountInBillCurr: allocatedInBillCurr,',
        '            allocatedAmountInBillingCurr: allocatedInBillingCurr,',
        '            exchangeRate:              bill.exchangeRate,',
        '        }',
        '        billAllocations.append(allocation)',
        '        remainingToAllocate -= abs(allocatedInBillingCurr)',
        '',
        '    return WriteOffTask{',
        '        taskId:           generateTaskId(),',
        '        invoiceNo:        invoice.invoiceNo,',
        '        writeOffAmount:   writeOffAmount,',
        '        writeOffCurrency: billingCurrency,',
        '        writeOffDate:     today(),',
        '        billAllocations:  billAllocations,',
        '    }',
        '}',
      ]),

      heading('4.2 回退未完全消耗的账单', HeadingLevel.HEADING_2),
      codeBlock([
        'function rollbackUnconsumedBills(',
        '    pendingBills:     List<PendingBill>,',
        '    accumulatedBills: List<PendingBill>,',
        '    writeOffAmount:   decimal,',
        '    currentIndex:     int',
        '): void {',
        '',
        '    remainingWriteOff = abs(writeOffAmount)',
        '    sign = writeOffAmount > 0 ? 1 : -1',
        '',
        '    // 从后往前回退（最后一个账单可能未完全消耗）',
        '    for i = accumulatedBills.length - 1 down to 0:',
        '        bill = accumulatedBills[i]',
        '        billAbs = abs(bill.remainingBillableAmount)',
        '',
        '        if billAbs <= remainingWriteOff:',
        '            // 该账单被完全消耗',
        '            remainingWriteOff -= billAbs',
        '            bill.remainingBillableAmount = 0',
        '        else:',
        '            // 该账单被部分消耗，回退剩余金额',
        '            consumed = remainingWriteOff',
        '            bill.remainingBillableAmount -= sign * consumed',
        '            remainingWriteOff = 0',
        '            break  // 该账单未消耗完，指针需回退到此位置',
        '',
        '    // 注：实际实现中需将 billIndex 回退到未消耗账单的位置',
        '}',
      ]),

      heading('4.3 取绝对值较小者', HeadingLevel.HEADING_2),
      codeBlock([
        'function minByAbs(a: decimal, b: decimal): decimal {',
        '    if abs(a) <= abs(b):',
        '        return a',
        '    else:',
        '        return b',
        '}',
      ]),

      // ========== 五、流程图解 ==========
      heading('五、算法核心流程图解', HeadingLevel.HEADING_1),

      boldPara('阶段一：数据准备（步骤1-4）'),
      codeBlock([
        '┌──────────────────────────────────────────────────────────────────┐',
        '│                     步骤1-4：数据准备                             │',
        '│                                                                  │',
        '│  ┌─────────────┐    ┌─────────────┐    ┌─────────────┐          │',
        '│  │ 客户开票配置 │───▶│ 自动核销    │───▶│ 待核销发票  │          │',
        '│  │（过滤条件）  │    │ 客户列表    │    │ 列表(FIFO)  │          │',
        '│  └─────────────┘    └─────────────┘    └──────┬──────┘          │',
        '│                                               │                  │',
        '│                     ┌─────────────┐    ┌──────▼──────┐          │',
        '│                     │ 账单(FIFO)  │◀───│ 待核销账单  │          │',
        '│                     │ 按账务日期  │    │ 列表(已转    │          │',
        '│                     └─────────────┘    │ 开票币种)   │          │',
        '│                                        └─────────────┘          │',
        '└──────────────────────────────────────────────────────────────────┘',
      ]),

      boldPara('阶段二：核销匹配循环（步骤5-6）'),
      codeBlock([
        '┌──────────────────────────────────────────────────────────────────┐',
        '│              步骤5-6：核销匹配主循环                              │',
        '│                                                                  │',
        '│  for each 发票 in 待核销发票列表:                                │',
        '│    ┌──────────────────────────────────────────┐                  │',
        '│    │ 发票金额 < 0 ?                            │                  │',
        '│    │                                           │                  │',
        '│    │ YES ──▶ 跳过正数账单                      │                  │',
        '│    │    └──▶ 累积负数账单直到汇总为负           │                  │',
        '│    │    └──▶ writeOff = min(|sum|, |invoice|)  │                  │',
        '│    │                                           │                  │',
        '│    │ NO  ──▶ 跳过负数/零账单                    │                  │',
        '│    │    └──▶ 累积正数账单                       │                  │',
        '│    │    └──▶ writeOff = min(sum, invoice)      │                  │',
        '│    └──────────────────────────────────────────┘                  │',
        '│    ┌──────────────────────────────────────────┐                  │',
        '│    │ 创建核销任务                               │                  │',
        '│    │ ├── 按比例分配金额到各账单                 │                  │',
        '│    │ ├── 转换为账单币种                         │                  │',
        '│    │ └── 持久化任务 + 更新发票/账单状态         │                  │',
        '│    └──────────────────────────────────────────┘                  │',
        '│    ┌──────────────────────────────────────────┐                  │',
        '│    │ 更新发票剩余可核销金额                     │                  │',
        '│    │ 扣减账单剩余金额                           │                  │',
        '│    │ 回退未消耗账单指针                         │                  │',
        '│    └──────────────────────────────────────────┘                  │',
        '│    while 发票剩余 != 0 AND 还有账单                              │',
        '└──────────────────────────────────────────────────────────────────┘',
      ]),

      // ========== 六、边界情况 ==========
      heading('六、边界情况处理', HeadingLevel.HEADING_1),

      createTable(
        ['场景', '处理方式'],
        [
          ['无符合条件的客户', '直接返回空列表，不创建任何任务'],
          ['待核销发票列表为空', '直接返回空列表'],
          ['待核销账单列表为空', '发票无法匹配，跳过该发票'],
          ['发票金额为0', '已在步骤2过滤（writeOffableAmount != 0）'],
          ['账单全部为正/负', '跳过无法匹配的账单，发票本轮无法核销'],
          ['发票类型限制', '按customerAccountCode关联同客户配置的账单'],
          ['汇率获取失败', '记录异常日志，跳过该账单'],
          ['汇率为0（除零保护）', 'exchangeRate非0判断，同币种使用1.0'],
        ]
      ),

      // ========== 附注 ==========
      new Paragraph({ spacing: { before: 400 }, children: [] }),
      new Paragraph({
        spacing: { before: 200 },
        children: [
          new TextRun({
            text: '—— 文档结束 ——',
            font: 'Microsoft YaHei',
            size: 21,
            color: '999999',
            italics: true
          })
        ]
      }),

    ]
  }]
});

// ===== Generate .docx =====
Packer.toBuffer(doc).then(buffer => {
  const outputPath = 'c:/Users/zt14033/WorkBuddy/2026-07-20-16-00-27/发票与账单自动核销算法设计.docx';
  fs.writeFileSync(outputPath, buffer);
  console.log('Document generated: ' + outputPath);
  console.log('Size: ' + (buffer.length / 1024).toFixed(1) + ' KB');
});
