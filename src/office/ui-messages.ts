// Text returned by Office.js and the native Office bridges. Input and document values are never translated.
export const officeUiMessages: Readonly<Record<string, string>> = {
  '当前 Office 无法读取准确页面尺寸，请下载 PPT 后手动导入。':
    'Office could not read the slide size. Download the deck and import the slides manually.',
  '演示稿超过 500 页，请拆分后再插入。':
    'The deck exceeds 500 slides. Split it before inserting slides.',
  '当前 PowerPoint 版本无法读取页面对象，请更新 Office。仍可根据素材生成并下载 PPT。':
    'This version of PowerPoint cannot read slide objects. Update Office, or generate and download a deck from your sources.',
  '当前版本无法读取选中页，请改为整稿范围。':
    'This version cannot read selected slides. Choose the entire deck.',
  '请在左侧缩略图中选中至少一页。': 'Select at least one slide thumbnail on the left.',
  '一次最多读取 60 页，演示稿最多 500 页。请缩小选区或拆分演示稿。':
    'Read up to 60 slides at a time from a deck of up to 500 slides. Select fewer slides or split the deck.',
  '目标页面已经删除，请重新读取。': 'The target slide was deleted. Read the slides again.',
  '所选页超过 1,500 个对象，请缩小范围。':
    'The selected slides contain more than 1,500 objects. Select fewer slides.',
  '所选页面内容过多，请缩小范围。':
    'The selected slides contain too much content. Select fewer slides.',
  '当前 PowerPoint 不支持页面重排，请更新 Office。':
    'This version of PowerPoint cannot reorder slides. Update Office.',
  '页面顺序核对未通过。': 'The slide order could not be verified.',
  '修改后的对象未找到。': 'The edited object could not be found.',
  '当前 Office 无法导出原稿，请使用“另存为”备份。':
    'Office cannot export the original presentation. Use Save As to make a backup.',
  '原稿超过 150 MB，请使用 Office“另存为”备份。':
    'The original presentation exceeds 150 MB. Use Save As in Office to back it up.',
  '原稿备份读取失败，请使用 Office“另存为”。':
    'The backup could not be read. Use Save As in Office.',
  '当前 Office 不支持插入整页，请下载后手动导入。':
    'This version of Office cannot insert complete slides. Download the deck and import the slides manually.',
  '演示稿页序或尺寸已变化，尚未插入。请检查当前演示稿并更新预览。':
    'The slide order or dimensions changed. Nothing was inserted. Check the presentation and refresh the preview.',
  '插入后的页数或位置核对未通过。': 'The slide count or insertion position could not be verified.',
  '插入未能确认全部成功，请检查演示稿末尾，必要时使用 Office 撤销。不要直接重复插入。':
    'Insertion could not be fully verified. Check the end of the deck before retrying. Use Undo in Office if necessary.',
  '无法确认选区结构。可生成和插入结果，请重新选择纯文字段落后再替换。':
    'The selection structure could not be verified. You can generate or insert a result. To replace text, select plain paragraphs again.',
  '选区含表格、图片、公式或其他文档对象。为保留结构，本次仅支持插入或复制；如需替换，请只选中纯文字段落。':
    'The selection includes a table, image, equation or another document object. Insert or copy the result to keep its structure, or select only plain paragraphs before replacing.',
  'Office.js 加载超时。请检查网络是否能访问 appsforoffice.microsoft.com，然后重试。':
    'Office.js timed out. Check access to appsforoffice.microsoft.com and try again.',
  'Office.js 加载失败，请检查网络或代理设置。':
    'Office.js could not load. Check the network or proxy settings.',
  '未连接到 Office。请从 Word、Excel 或 PowerPoint 的加载项入口打开。浏览器体验请移除地址中的 ?office=1。':
    'Open the add-in from Word, Excel or PowerPoint. For a browser demo, remove ?office=1 from the URL.',
  'Office 初始化失败，请重新打开加载项。': 'Office could not initialize. Reopen the add-in.',
  '当前宿主不受支持，请在 Word、Excel 或 PowerPoint 中打开此加载项。':
    'This host is unsupported. Open the add-in in Word, Excel or PowerPoint.',
  '目标内容或选区已经变化。请重新读取选区，再预览并确认写入。':
    'The target or selection changed. Read it again, then preview and confirm the write.',
  '没有读取到内容。请先在文档中选中文字或单元格。':
    'The selection is empty. Select text or cells in the document first.',
  '选区内容超过 30,000 字符，请缩小选区。':
    'The selection exceeds 30,000 characters. Select less content.',
  '无法读取选区。请进入文本编辑状态，选中文字后重试。':
    'The selection could not be read. Enter text-editing mode, select the text and try again.',
  'Office 无法写入当前选区。请检查文档是否只读，或复制结果手动粘贴。':
    'Office could not write to the selection. Check that the document is editable, or copy and paste the result.',
  '无法确认原段落样式，请重新选择写入位置。':
    'The original paragraph style could not be verified. Select the insertion point again.',
  '无法确认段落边界，请重新选择写入位置。':
    'The paragraph boundaries could not be verified. Select the insertion point again.',
  'Word · 已读取的文字选区': 'Word · selected text',
  '在已确认光标或选区后新起段落插入报告（保留原文）':
    'Insert the report in a new paragraph after the cursor or selection and keep the original.',
  '在已确认光标或选区后新起段落插入排版内容（保留原文）':
    'Insert formatted content in a new paragraph after the cursor or selection and keep the original.',
  '在已确认选区末尾插入（保留原文）': 'Insert after the selection and keep the original.',
  在已确认光标处插入: 'Insert at the confirmed cursor position.',
  '公式将在 Excel 中计算。创建失败时可能留下新工作表，请按错误提示检查后再试。':
    'Excel will calculate the formulas. If creation fails, check for a partially created sheet before retrying.',
  '数据表与确认时的内容不同，请返回检查后重新确认。':
    'The table changed after confirmation. Review and confirm it again.',
  '目标工作表名称已被使用，请返回检查后重新创建。':
    'The target sheet name is already in use. Review the name and create the sheet again.',
  '一次最多读取 2,000 个单元格，请缩小选区。':
    'Read up to 2,000 cells at a time. Select a smaller range.',
  '选区为空，请选中包含数据或公式的单元格。':
    'The selection is empty. Select cells containing data or formulas.',
  '请在 Excel 中选中一个目标单元格，再点击写入公式。':
    'Select one target cell in Excel before writing the formula.',
  '（空单元格）': '(empty cell)',
  '无法确认幻灯片尺寸，请复制大纲后手动排版。':
    'The slide dimensions could not be verified. Copy the outline and lay out the slides manually.',
  '建页期间演示文稿页序发生变化，已停止填充。请检查新增的空白页并撤销后重试。':
    'The slide order changed during creation. Check the new blank slides and undo if necessary before retrying.',
  '当前 Office 无法读取选区。可将内容粘贴到输入框。':
    'Office cannot read this selection. Paste the content into the reference field.',
  '当前 Word 版本不支持自动写入排版结果，请更新 Office，或复制结果手动粘贴。':
    'This version of Word cannot write formatted content automatically. Update Office or copy and paste the result.',
  '将按 Markdown 排版替换选区，原选区的字体和段落样式可能改变。':
    'Replacing the selection applies Markdown formatting and may change the original fonts and paragraph styles.',
  '将以纯文本替换，选区内部的字体、加粗和段落样式可能改变。需要保留排版时，请返回并选择插入。':
    'Plain text replacement may change fonts, bold text and paragraph styles. Choose Insert to preserve the original formatting.',
  在演示文稿末尾添加幻灯片: 'Append slides to the end of the presentation',
  保留所有现有幻灯片: 'Keep all existing slides',
  '此版本 Office 不支持该写入操作。请复制结果并手动粘贴。':
    'This version of Office cannot perform this write. Copy and paste the result manually.',
  'Office 操作失败，可能是目标已改变或文档只读。请重新选择内容后重试；结果仍可复制。':
    'Office could not complete the change. Check that the target still exists and the document is editable. You can still copy the result.',
  '操作未完成，请重试。': 'The operation did not finish. Try again.',
  '本地工作台无法识别这项模型操作。': 'The local workbench does not recognize this model action.',
  'Office 未完成操作，请检查当前文档。':
    'Office did not complete the operation. Check the current document.',
  '替换会改变选区内的文字和格式，请核对预览。':
    'Replacement will change text and formatting inside the selection. Review the preview.',
  '此入口无法预览页面缩略图。': 'This entry point cannot preview slide thumbnails.',
  '请在 PowerPoint 中另存为副本。': 'Use Save As in PowerPoint to make a copy.',
  '无效的 Office 应用。': 'Invalid Office application.',
  '侧边栏已关闭。': 'The task pane is closed.',
  '未知的侧边栏请求。': 'Unknown task pane request.',
  '请求繁忙，请重试。': 'Another request is running. Try again.',
  'Office 未能完成该操作。': 'Office could not complete the operation.',
  'Office 未响应。请在目标应用打开文档，并允许系统请求的自动化访问。':
    'Office did not respond. Open a document in the target app and allow automation access when prompted.',
  'Office 未完成操作。': 'Office did not complete the operation.',
  '工作台已关闭。': 'The workbench is closed.',
  '上一项 Office 操作仍在进行。': 'The previous Office operation is still running.',
  '当前应用不支持这项写入。': 'This app does not support that write operation.',
  '写入位置已过期，请重新读取。': 'The insertion point is stale. Read the target again.',
  '写入内容无效或过大。': 'The content is invalid or too large to write.',
  '幻灯片改动无效。': 'The proposed slide changes are invalid.',
  '待插入页数无效。': 'The number of slides to insert is invalid.',
  '演示稿文件过大。': 'The presentation file is too large.',
  '演示稿格式无效。': 'Invalid presentation file.',
  '已取消保存演示稿。': 'Saving the presentation was cancelled.',
  '演示稿已保存。请在 PowerPoint 中打开刚保存的文件。':
    'Presentation saved. Open the saved file in PowerPoint.',
  '此操作不适用于当前文档。': 'This operation is unavailable for the current document.',
  '无效的窗口。': 'Invalid window.',
  '窗口已关闭。': 'The window is closed.',
  '文档、选区或内容已变化，请重新读取。':
    'The document, selection or content changed. Read it again.',
  '选区过大，请缩小范围。': 'The selection is too large. Select less content.',
  '格式化内容文件已失效，请重新生成。':
    'The temporary formatted file is no longer available. Generate the result again.',
  '请在普通文字段落中插入格式化内容。':
    'Select a plain text paragraph for inserting formatted content.',
  '无法确认格式化内容的段落边界。':
    'The boundaries of the formatted paragraph could not be verified.',
  '格式化内容为空，请重新生成。': 'The formatted content is empty. Generate the result again.',
  '目标文档已改变，请重新读取并确认写入位置。':
    'The target document changed. Read and confirm the insertion point again.',
  '文档为只读，请打开可编辑副本。': 'The document is read-only. Open an editable copy.',
  '请只选择普通文字。': 'Select plain text only.',
  '写入位置的段落边界已改变，请重新读取并确认写入位置。':
    'The paragraph boundaries at the insertion point changed. Read and confirm the target again.',
  '请在 Word 中先选中文字。': 'Select text in Word first.',
  'Word · 当前选区': 'Word · current selection',
  '选区包含文档对象，可保留原文并在选区后插入结果。':
    'The selection includes document objects. Insert after the selection to keep the original.',
  'Word · 在当前选区后插入': 'Word · insert after the current selection',
  '未知 Word 操作。': 'Unknown Word operation.',
  '新建工作表，保留已有数据': 'Create a worksheet and keep the existing data',
  'Excel · 新建数据表': 'Excel · create a data table',
  '工作簿为只读。': 'The workbook is read-only.',
  '数据表尺寸无效。': 'Invalid table dimensions.',
  '数据表的列数不一致。': 'The rows have different numbers of columns.',
  '请选择不超过 2,000 个单元格的一块连续区域。':
    'Select one continuous range of up to 2,000 cells.',
  '请选择可编辑的普通单元格。': 'Select a regular, editable cell.',
  '演示稿超过 500 页。': 'The presentation exceeds 500 slides.',
  '每次最多编辑 60 页。': 'Edit up to 60 slides at a time.',
  '选中的对象过多，请缩小范围。': 'Too many objects selected. Reduce the selection.',
  '所选页面没有可读取的文字。': 'The selected slides contain no readable text.',
  'PowerPoint · 所选页面': 'PowerPoint · selected slides',
  保留现有页面: 'Keep the existing slides',
  'PowerPoint · 在末尾添加页面': 'PowerPoint · append slides',
  '导入页数与预览不一致。': 'The number of imported slides differs from the preview.',
  '未知 PowerPoint 操作。': 'Unknown PowerPoint operation.',
  '对象已移动。': 'The object has moved.',
  '请在 PowerPoint 中撤销部分改动并检查目标页面。':
    'Inspect the target slides and undo partial changes in PowerPoint as needed.',
  'Office 窗口已关闭，请重新打开侧边栏。': 'The Office window was closed. Reopen the task pane.',
  '当前 Office 窗口与此侧边栏不一致，请回到原窗口后重试。':
    'This task pane belongs to a different Office window. Return to its original window and try again.',
  '未知 Office 应用。': 'Unknown Office app.',
  'Office 应用与侧边栏不一致。': 'The Office application does not match this task pane.',
  '无法完成 Office 操作。请打开可编辑文档，检查当前选区，然后重试。':
    'Office could not complete the operation. Open an editable document, check the selection and try again.',
  '侧边栏已关闭，请重新打开。': 'The task pane is closed. Reopen it.',
  '无法绑定 Office 窗口。': 'Could not connect to the Office window.',
  '侧边栏已绑定另一个 Office 窗口。':
    'The task pane is already connected to another Office window.',
  'Office 应用与窗口不一致。': 'The Office application and window do not match.',
  '侧边栏组件缺失，请关闭 Office 后重新运行安装程序。':
    'A task pane component is missing. Close Office and run the installer again.',
  '请从 Office 窗口打开侧边栏。': 'Open the task pane from the Office window.',
  '不支持当前 Office 进程架构。': 'This Office process architecture is unsupported.',
  '侧边栏文件缺失，请重新解压安装包并运行安装程序。':
    'Task pane files are missing. Extract the entire package and run the installer again.',
  '侧边栏组件缺失，请重新解压安装包并运行安装程序。':
    'A task pane component is missing. Extract the entire package and run the installer again.',
  '无法加载侧边栏，请关闭并重新打开。': 'The task pane could not load. Close and reopen it.',
  '侧边栏浏览器已停止，请关闭并重新打开。': 'The task pane browser stopped. Close and reopen it.',
  '无法启动侧边栏，请关闭 Office 后重新运行安装程序。':
    'The task pane could not start. Close Office and run the installer again.',
  '侧边栏已断开，请关闭并重新打开。': 'The task pane disconnected. Close and reopen it.',
  '侧边栏未能完成操作。': 'The task pane could not complete the operation.',
  '请求繁忙，请稍后重试。': 'Another request is running. Try again shortly.',
  '工作台文件缺失，请重新解压安装包并运行安装程序。':
    'Workbench files are missing. Extract the entire installer package and run setup again.',
  '无法启动工作台后台进程。': 'The workbench background process could not start.',
  '工作台后台进程已断开，请关闭并重新打开侧边栏。':
    'The workbench background process disconnected. Close and reopen the task pane.',
  '无效的 Office 操作。': 'Invalid Office operation.',
  '未知的侧边栏操作。': 'Unknown task pane operation.',
  '文档、选区或内容已发生变化。请重新读取并确认写入位置。':
    'The document, selection or content changed. Read the target again and confirm the insertion point.',
  '选区文字过长，请缩小选区再读取。':
    'The selected text is too long. Select a shorter passage and read it again.',
  '无法唯一定位目标 Word 文档，请关闭重复副本后重新读取。':
    'Several open copies match the target document. Close the duplicates and read it again.',
  '无法唯一定位目标 Word 文档，请关闭同名副本后重新读取。':
    'Several open documents have the same name. Close the duplicates and read the intended document again.',
  '请在 Word 中先选中需要处理的文字。': 'Select the text to work on in Word first.',
  '选区包含表格、域或其他文档对象，可以读取并在选区后插入结果。':
    'The selection includes tables, fields or other objects. Read the content and insert after it to keep the original.',
  '这段选区包含特殊对象，请使用插入。': 'The selection contains special objects. Choose Insert.',
  '无法确认此处的文字范围，请在正文中插入格式化内容。':
    'The text range could not be verified. Insert the formatted content in the document body.',
  'Word 无法定位这份文档的文字范围。请保存并重新打开该文档，再重新读取选区。':
    'Word could not locate this document’s text range. Save and reopen the document, then read the selection again.',
  '无法确认选区的文字范围，请重新读取并确认写入位置。':
    'The text range could not be verified. Read and confirm the insertion point again.',
  '格式化临时文档已打开，请重新生成。':
    'The temporary formatted document is already open. Generate the result again.',
  '无法唯一定位格式化临时文档，请重新生成。':
    'The temporary formatted document could not be identified. Generate the result again.',
  '未能确认选区末尾的插入位置，请重新选择。':
    'The end of the selection could not be verified. Select the insertion point again.',
  'Excel · 新建工作表': 'Excel · create a worksheet',
  '数据表尺寸超出限制。': 'The table exceeds the supported dimensions.',
  '请把选区缩小到 2,000 个单元格以内。': 'Select a range of up to 2,000 cells.',
  '无法确认选区大小，请选择一块连续的单元格区域。':
    'The range size could not be verified. Select one continuous range.',
  '请选择一个公式目标单元格。': 'Select one cell for the formula.',
  '请选择普通单元格，不能覆盖合并区域或数组公式。':
    'Select a regular cell. Merged cells and array formulas cannot be overwritten.',
  '演示稿超过 500 页，请拆分后再编辑。': 'The deck exceeds 500 slides. Split it before editing.',
  '每次最多处理 60 页，请在 PowerPoint 中缩小选区。':
    'Process up to 60 slides at a time. Select fewer slides in PowerPoint.',
  '无法识别所选页面。': 'The selected slides could not be identified.',
  '选区对象过多，请减少选中的页面。':
    'The selection contains too many objects. Select fewer slides.',
  '选中的页面没有可读取的文字。': 'The selected slides contain no readable text.',
  'PowerPoint · 选中页面': 'PowerPoint · selected slides',
  'PowerPoint · 末尾添加页面': 'PowerPoint · append slides',
  '该操作不适用于此 PowerPoint 入口。':
    'This operation is unavailable through this PowerPoint entry point.',
  '请在 PowerPoint 中调整页面顺序。': 'Reorder the slides in PowerPoint.',
  '页面对象已变化。': 'The slide objects changed.',
  '部分改动需要在 PowerPoint 中撤销，请检查目标页面。':
    'Inspect the target slides and undo any partial changes in PowerPoint.',
  '未知的 Office 应用。': 'Unknown Office application.',
  '请在“系统设置 → 隐私与安全性 → 自动化”允许工作台控制目标 Office 应用。':
    'In System Settings > Privacy & Security > Automation, allow the workbench to control the target Office app.',
  'Office 未完成操作。请检查当前文档、选区及编辑权限。':
    'Office did not finish the operation. Check the document, selection and editing permissions.',
};
