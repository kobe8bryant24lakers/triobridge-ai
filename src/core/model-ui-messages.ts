// Presentation-only translations of diagnostics from the model service, storage and output validators.
export const modelUiMessages: Readonly<Record<string, string>> = {
  '结果不是有效的单行 Excel 公式。请编辑结果为以 = 开头的公式后再写入。':
    'Enter a single Excel formula starting with = before writing it.',
  '公式的方括号不匹配，请检查表格列引用。':
    'The square brackets do not match. Check the table column references.',
  '公式的括号不匹配，请检查后再写入。':
    'The parentheses do not match. Check the formula before writing it.',
  '公式的引号或括号未闭合，请检查后再写入。':
    'A quotation mark or parenthesis is unclosed. Check the formula before writing it.',
  '公式含外部链接、工作簿引用或外部调用，请移除后重试。':
    'Remove external links, workbook references or external calls from the formula, then try again.',
  '大纲格式无效，需要 {"slides":[{"title":"标题","bullets":["要点"]}]}。可编辑结果或重新生成。':
    'The outline must use {"slides":[{"title":"Title","bullets":["Point"]}]}. Edit the result or generate it again.',
  '大纲包含暂不支持的字段，请仅保留 slides，并将文字放入页面标题或要点。':
    'The outline has unsupported fields. Keep only slides and put the content in slide titles or bullets.',
  '大纲需要包含 1–12 页。': 'An outline must contain 1–12 slides.',

  '模型配置格式无效。': 'The model settings format is invalid.',
  '模型配置缺少必要字段。': 'The model settings are missing required fields.',
  '本地配置无法读取。原始数据未被覆盖，可在模型设置中重新保存配置。':
    'The saved settings could not be read. The original data is intact. Try saving the profile again in Model Settings.',
  '最多保存 30 个模型配置。': 'You can save up to 30 model profiles.',
  '无法保存配置，浏览器存储可能被禁用或已满。请允许本地存储后重试。':
    'The settings could not be saved. Browser storage may be blocked or full. Enable local storage or free space, then try again.',
  '浏览器存储不可用。': 'Browser storage is unavailable.',
  '无法读取本机模型配置。请核对当前账户的文件权限并点击重新连接。':
    'The local model settings could not be read. Check the current account’s file permissions, then reconnect.',
  '本机共享入口缺少有效的安装凭证。请重新运行本机共享安装包中的 Install，保存文档并完全退出、重开 Office。旧模型配置未修改。':
    'This entry point has no valid installation credential. Run Install from the local shared package, save your documents, then fully quit and reopen Office. Your previous model settings are intact.',
  '本机共享服务暂时未连接。请双击本机共享安装包中的 Restart，恢复后重试。已保留共享配置，未切换到旧本地密钥。':
    'The local shared service is unavailable. Run Restart from the local shared package, then try again. Shared settings have been retained, and legacy credentials are not being used.',
  '正在保存模型配置，请等待当前操作完成。':
    'A model profile is being saved. Wait for it to finish.',
  '请重新读取共享配置后再修改。': 'Reload the shared settings before editing.',
  '请选择仍存在的旧模型配置后再迁移。': 'Choose existing legacy model profiles to migrate.',
  '旧配置在迁移期间发生变化，已保留本地副本，请核对后手动清理。':
    'The legacy settings changed during migration. A local copy has been kept. Review it before removing it.',
  '共享迁移已成功，但旧浏览器副本未能清理，请核对后清理该应用的旧配置。':
    'Migration succeeded, but the old browser copy could not be removed. Review and remove the legacy settings in this app.',
  '共享模式不能写入旧浏览器配置。': 'Shared mode cannot write to the legacy browser settings.',
  '正在连接共享配置，请稍候。': 'Connecting to shared settings. Please wait.',
  '配置文件过大。': 'The configuration file is too large.',
  '请选择有效的 JSON 配置文件。': 'Choose a valid JSON settings file.',
  '配置文件需包含 1–30 个模型，格式版本为 1。':
    'The settings file must use version 1 and contain 1–30 model profiles.',

  '请填写服务名称和模型 ID。': 'Enter a service name and a model ID.',
  '模型配置字段过长，请检查输入。':
    'One or more model settings exceed the allowed length. Check the fields.',
  '共享模型版本无效，请刷新配置后重试。':
    'The shared model revision is invalid. Refresh the settings and try again.',
  '超时应为 5–300 秒。': 'Set a timeout of 5–300 seconds.',
  '服务拒绝了请求，请检查模型 ID 和接口参数。':
    'The service rejected the request. Check the model ID and API parameters.',
  'API Key 无效或已过期，请在模型设置中更新。':
    'The API key is invalid or expired. Update it in Model Settings.',
  '当前 API Key 没有访问权限，请检查模型权限。':
    'The API key does not have access to this model. Check its permissions.',
  '找不到接口或模型，请检查 Base URL、接口格式和模型 ID。':
    'The endpoint or model was not found. Check the base URL, API format and model ID.',
  '请求过于频繁或额度不足，请稍后重试并检查服务额度。':
    'The request was rate-limited or the quota was exhausted. Check the service quota and try again later.',
  '已取消生成。': 'Generation canceled.',
  '服务返回的数据格式无效，请检查所选接口格式。':
    'The service returned an invalid response. Check the API format.',
  '模型输出被长度限制截断，请缩短任务或调整服务端输出上限后重试。':
    'The model output reached the service’s length limit. Shorten the task or adjust the server’s output limit and try again.',
  '模型上下文空间已用尽，输出被截断。请减少输入内容，或在服务端增大上下文窗口后重试。':
    'The model ran out of context space and stopped early. Reduce the input or increase the server’s context window and try again.',
  '模型未完成此请求，请调整输入后重试。':
    'The model did not complete this request. Revise the input and try again.',
  '服务未返回可用文本，请检查所选接口格式。':
    'The service returned no usable text. Check the API format.',
  '服务返回了无效 JSON，请检查接口地址。':
    'The service returned invalid JSON. Check the endpoint URL.',
  '模型请求未完成。请重新打开工作台，在模型设置中测试该 API 的连接。':
    'The model request did not complete. Reopen the workspace and test this API in Model Settings.',
  '无法连接 TrioBridge 本机共享服务。请双击本机共享安装包中的 Restart，重启服务后重新连接。':
    'The TrioBridge local shared service could not be reached. Run Restart from the local shared package and reconnect.',
  '当前加载项正在直连本地 API，连接失败。请打开「三桥 AI（本机共享）」并确认模型设置显示「本机共享」；共享版通过本机服务调用，不要求 API 开放浏览器 CORS。':
    'The add-in could not connect directly to the local API. Open the locally shared TrioBridge AI add-in and check that Model Settings shows “Shared on this device”. The shared service calls the API without requiring browser CORS access.',
  '无法连接模型服务。请检查地址、网络和 CORS；本地 HTTP 模型可配置本机代理。':
    'The model service could not be reached. Check the address, network and CORS settings. A local proxy can be configured for a local HTTP model.',

  '读取的共享模型配置无效，请重新加载。': 'The shared model settings are invalid. Reload them.',
  '请从本机安装的共享加载项打开。当前入口没有本机共享凭据。':
    'Open the locally installed shared add-in. This entry point has no credential for the local service.',
  '暂时无法读取共享配置，请重新打开模型设置后重试。':
    'Shared settings could not be read. Reopen Model Settings and try again.',
  '共享模型操作未完成，请重试。': 'The shared model operation did not complete. Try again.',
  '已取消操作。': 'Operation canceled.',
  '无法从此入口读取共享配置。': 'Shared settings cannot be read from this entry point.',
  '连接超时。请查看系统是否在请求密钥访问，并检查模型 API 是否可用。重新连接后可核对配置。':
    'The connection timed out. Check for a system credential prompt and verify that the model API is available. Reconnect and review the settings.',
  '读取或保存共享模型配置超时，请重试。': 'Reading or saving shared settings timed out. Try again.',
  '工作台未能读取或保存共享配置。请重新打开模型设置，核对后重试。':
    'The workspace could not read or save shared settings. Reopen Model Settings, review them and try again.',
  '无法连接 TrioBridge 本机共享服务。请双击本机共享安装包中的 Restart，重启后重试。本次修改尚未确认保存，请重新连接后核对配置。':
    'The TrioBridge local shared service could not be reached. Run Restart from the local shared package. Reconnect and review the settings to check whether the change was saved.',

  '结构化结果超过 100,000 字符，请缩小任务范围。':
    'The structured result exceeds 100,000 characters. Reduce the scope of the task.',
  '结果需要有效的 JSON 对象。可编辑修正，或重新生成。':
    'The result must be a valid JSON object. Edit it or generate it again.',
  '报告需要 1–12 个章节。': 'A report must contain 1–12 sections.',
  '工作表名称不能包含 \\ / * ? : [ ]、控制字符，或以单引号开头结尾。':
    'A worksheet name cannot contain \\ / * ? : [ ], control characters or leading or trailing apostrophes.',
  '数据表需要 1–200 行，含表头最多 2,000 个单元格。':
    'A data table must contain 1–200 data rows and at most 2,000 cells including headers.',
  '列格式需要与表头一一对应，可选 general、text、integer、number、percent。':
    'Specify one format per column: general, text, integer, number or percent.',

  '报告结构无效。': 'The report structure is invalid.',

  '请求过大或操作无效。': 'The request is too large or the operation is invalid.',
  '无效的模型请求。': 'The model request is invalid.',
  '模型操作包含不支持的字段。': 'The model operation contains unsupported fields.',
  '无法读取当前账户的共享配置，请检查配置目录权限。':
    'Shared settings could not be read for the current account. Check the settings folder permissions.',

  '共享配置文件暂时被占用。请待占用释放后重试。':
    'The shared settings file is in use. Try again after it is released.',
  '当前账号无法访问共享配置目录。': 'The current account cannot access the shared settings folder.',
  '后台服务无法读取本次安装的共享访问凭据。请用平时打开 Office 的账号运行 Install。':
    'The background service cannot read the shared credential for this installation. Run Install using the account you normally use for Office.',
  '后台服务无法检查共享配置的 Windows 访问权限。':
    'The background service could not check the Windows permissions for the shared settings.',
  '共享文件的所有者与运行服务的账号不匹配。请使用原安装账号。':
    'The owner of the shared file differs from the account running the service. Use the account that installed it.',
  '后台服务无法设置共享文件的专用权限。请检查当前账号的目录权限。':
    'The background service could not set private permissions on the shared file. Check the current account’s folder permissions.',
  '共享文件权限校验未通过。请检查现有文件所有者和访问权限。':
    'Shared file permissions could not be verified. Check the file owner and access permissions.',
  '后台服务无法读取 profiles.json。请检查当前账号的文件访问权限。':
    'The background service cannot read profiles.json. Check the file permissions for the current account.',
  'profiles.json 的文件类型或所有者不符合要求。请检查原文件。':
    'The file type or owner of profiles.json is unexpected. Check the original file.',
  'profiles.json 超出共享配置大小限制。请保留原文件并检查备份。':
    'The size of profiles.json exceeds the shared settings limit. Keep the original file and check your backups.',
  'profiles.json 无法解析为 JSON。请保留原文件并核对可信备份。':
    'The contents of profiles.json are not valid JSON. Keep the original and check a trusted backup.',
  'profiles.json 的结构或版本无法识别。请保留原文件并核对可信备份。':
    'The structure or version of profiles.json is unrecognized. Keep the original and check a trusted backup.',
  '后台服务无法初始化共享配置组件。':
    'The background service could not initialize shared settings.',

  '只允许同源的本机加载项访问共享模型。':
    'Only a local add-in from the same origin can access shared models.',
  '此加载项尚未连接本机共享服务，请使用本机安装入口重新打开。':
    'This add-in is not connected to the local shared service. Reopen it using the locally installed entry point.',
  '共享模型仅接受 JSON POST。': 'The shared model service accepts only JSON POST requests.',
  '共享模型接口不存在。': 'The shared model endpoint does not exist.',
  '共享模型请求过大。': 'The shared model request is too large.',
  '请求需要有效的 JSON。': 'The request must contain valid JSON.',
  '请求包含不支持的字段。': 'The request contains unsupported fields.',
  '本机共享模型服务暂时不可用，未降级保存密钥。':
    'The local shared model service is unavailable. Keys have not been saved to an alternative store.',

  '本机共享配置暂时不可用，请检查本机服务与安装。':
    'Local shared settings are unavailable. Check the local service and installation.',

  '本机代理尚未配置 MODEL_PROXY_TARGET。': 'MODEL_PROXY_TARGET is not set for the local proxy.',
  '本机代理目标必须是服务根地址。': 'The local proxy target must be a service root URL.',
  '模型配置包含不支持的字段。': 'The model settings contain unsupported fields.',
  '请检查模型名称、ID 和 5–300 秒的超时设置。':
    'Check the model name, ID and timeout (5–300 seconds).',
  '请选择保留、替换或删除密钥。': 'Choose whether to keep, replace or remove the key.',
  '请输入有效的新 API Key。': 'Enter a valid new API key.',
  '只有替换密钥时可以提交 API Key。': 'An API key may be submitted only when replacing the key.',
  '模型请求需要 1–16 条消息。': 'A model request requires 1–16 messages.',
  '模型消息格式无效。': 'The model message format is invalid.',
  '模型消息内容无效。': 'The model message content is invalid.',
  '只支持本次选择的 PNG / JPEG 图片数据，不支持远程图片地址。':
    'Only PNG or JPEG data from the selected images is supported. Remote image URLs are unsupported.',
  '图片过多或过大，请减少图片数量。':
    'There are too many images or they are too large. Select fewer images.',
  '模型请求内容过长。': 'The model request content is too long.',
  '模型服务返回了无效 JSON。': 'The model service returned invalid JSON.',
  '模型服务拒绝了请求，请检查接口格式和模型 ID。':
    'The model service rejected the request. Check the API format and model ID.',
  'API Key 无效或已过期。': 'The API key is invalid or expired.',
  'API Key 没有此模型的访问权限。': 'The API key does not have access to this model.',
  '找不到模型接口或模型，请检查地址、接口格式和模型 ID。':
    'The model or its endpoint was not found. Check the address, API format and model ID.',
  '请求过于频繁或额度不足，请稍后重试。':
    'The request was rate-limited or the quota was exhausted. Try again later.',
  '模型服务未返回兼容的文本响应，请检查接口格式。':
    'The model service did not return a compatible text response. Check the API format.',
  '配置已更新，部分旧密钥尚待从系统安全存储清理；请重试清理。':
    'The settings were updated. Some old keys still need to be removed from the system credential store. Retry cleanup.',
  '无法保存共享配置，请检查本机目录权限和可用空间。':
    'Shared settings could not be saved. Check local folder permissions and available space.',
  '无法锁定共享配置，请检查目录权限。':
    'Shared settings could not be locked. Check the folder permissions.',
  '共享配置锁文件不安全。': 'The shared settings lock file failed its security check.',
  '其他 Office 正在保存模型配置，请刷新后重试。':
    'Another Office app is saving model settings. Refresh and try again.',
  '共享配置正在更新，请重试。': 'Shared settings are being updated. Try again.',
  '缺少有效的配置版本。': 'A valid settings revision is required.',
  '模型配置已在另一 Office 中更改。请刷新后重新确认本次修改。':
    'Another Office app changed this profile. Refresh and confirm your changes again.',
  '系统安全存储不可用或未解锁，密钥没有降级保存到文件。':
    'The system credential store is unavailable or locked. No key was saved to a file.',
  '此模型的系统密钥已缺失，请重新填写 API Key。':
    'The key for this model is missing from the system credential store. Enter the API key again.',
  '更换服务地址或接口格式后，请重新填写密钥或选择不使用密钥。':
    'After changing the service address or API format, enter the key again or choose No Key.',
  '已取消保存。': 'Save canceled.',
  '无法保存到系统安全存储，本次配置未生效。':
    'The system credential store could not save the change. These settings were not applied.',
  '不支持此界面语言。': 'This interface language is unsupported.',
  '一次可导入 1–30 个模型。': 'You can import 1–30 model profiles at a time.',
  '导入内容包含重复模型 ID。': 'The import contains duplicate model IDs.',
  '共享模型总数不能超过 30 个。': 'There can be no more than 30 shared models.',
  '导入包含已有模型 ID，不会覆盖已有配置。':
    'The import contains an existing model ID. Existing settings have not been overwritten.',
  '导入后的默认模型不存在。': 'The default model specified in the import does not exist.',
  '模型已被删除或不存在。': 'This model was deleted or does not exist.',
  '模型已被删除或不存在，请刷新模型列表。':
    'This model was deleted or does not exist. Refresh the model list.',
  '此模型已在另一 Office 中更新，请刷新后再生成。':
    'Another Office app updated this model. Refresh before generating again.',

  '请选择支持的接口格式（OpenAI 或 Claude）。': 'Select a supported API format: OpenAI or Claude.',
  '模型服务地址无效。': 'The model service URL is invalid.',
  '请填写完整的 API Base URL，例如 http://127.0.0.1:8001/v1。':
    'Enter a complete API base URL, such as http://127.0.0.1:8001/v1.',
  '服务地址须使用 HTTP(S)，且不能包含账号、密码、查询参数或片段。':
    'The service URL must use HTTP or HTTPS and cannot contain a username, password, query string or fragment.',
  '远程模型服务需要 HTTPS，HTTP 仅限本机回环地址。':
    'Remote model services require HTTPS. HTTP is allowed only for local loopback addresses.',
  '该地址是 OpenAI Chat Completions 接口。请选择 OpenAI 格式，或填写 Claude 接口地址。':
    'This is an OpenAI Chat Completions endpoint. Choose the OpenAI format or enter a Claude endpoint URL.',
  '连接测试未返回兼容的文本响应，请检查接口格式。':
    'The connection test did not return a compatible text response. Check the API format.',
  '模型未完成连接测试，请检查模型权限或服务策略。':
    'The model did not complete the connection test. Check model permissions or the service policy.',
  '模型返回内容过长，请检查服务端配置。':
    'The model returned too much content. Check the server settings.',
  '测试回复达到服务端输出长度上限；正式生成前请提高服务端输出上限。':
    'The test response reached the server’s output limit. Increase the server’s output limit before generating.',
  '测试回复因模型上下文空间用尽而停止；正式生成前请减少输入内容，或在服务端增大上下文窗口。':
    'The test response used up the model’s context window. Reduce the input or increase the server’s context window before generating.',
  '连接测试未返回可用文本，请检查模型输出和接口格式。':
    'The connection test returned no usable text. Check the model output and API format.',
  'Claude 图片输入需要本地 PNG 或 JPEG 图片。':
    'Claude image inputs require local PNG or JPEG images.',
  'Claude 系统指令需要文本。': 'Claude system instructions must be text.',
  'Claude 请求至少需要一条用户消息。': 'A Claude request requires at least one user message.',
  '模型信息接口未返回 JSON。': 'The model information endpoint did not return JSON.',
  '模型信息接口返回的数据过大。': 'The model information response is too large.',
  '无法从当前 Claude 接口地址查询模型输出上限。':
    'The model’s output limit cannot be queried from this Claude endpoint.',
  'Claude 接口要求 max_tokens，但模型信息未提供有效的 max_tokens 输出上限。请让服务提供该模型的真实上限后重试；若同一服务提供 OpenAI 兼容接口，也可在模型设置中手动选择该格式与地址。':
    'The Claude endpoint requires max_tokens, but the model metadata does not specify a valid output limit. Configure the service to report the model’s actual limit and try again. If the same service offers an OpenAI-compatible API, you can select that format and address in Model Settings.',
  'Claude 接口要求 max_tokens，但无法读取同一服务的模型输出上限。请检查模型信息接口、网络和访问权限。':
    'The Claude endpoint requires max_tokens, but the model’s output limit could not be read from the service. Check the model information endpoint, network and permissions.',
  '服务未返回 Claude Messages 响应，请检查接口格式和地址。':
    'The service did not return a Claude Messages response. Check the API format and address.',
  'Claude 接口未返回完整的最终回复，请检查模型输出和接口格式。':
    'The Claude endpoint did not return a complete final response. Check the model output and API format.',
  'Claude 接口返回了无效的文本片段。': 'The Claude endpoint returned an invalid text segment.',
  'Claude 接口未返回可用的文本内容。': 'The Claude endpoint returned no usable text.',
};

type MessagePattern = readonly [RegExp, string];

const variableMessages: readonly MessagePattern[] = [
  [
    /^第 (\d+) 页包含暂不支持的字段，请将内容移入标题或要点。$/,
    'Slide $1 contains unsupported fields. Move their content into the title or bullet points.',
  ],
  [
    /^第 (\d+) 页格式有误：需要一个不超过 60 字符的标题，以及 1–5 条不超过 160 字符的要点。$/,
    'Slide $1 needs a title of up to 60 characters and 1–5 bullet points of up to 160 characters each.',
  ],
  [
    /^模型服务返回 HTTP (\d{3})，请稍后重试。$/,
    'The model service returned HTTP $1. Try again later.',
  ],
  [/^模型服务返回 HTTP (\d{3})。$/, 'The model service returned HTTP $1.'],
  [
    /^请求超过 (\d+(?:\.\d+)?) 秒，请重试或增加超时时间。$/,
    'The request exceeded $1 seconds. Try again or increase the request timeout.',
  ],
  [
    /^Claude 接口要求 max_tokens，但模型信息查询返回 HTTP (\d{3})。请确认服务支持查询该模型的输出上限；若同一服务提供 OpenAI 兼容接口，也可在模型设置中手动选择该格式与地址。$/,
    'The Claude endpoint requires max_tokens. Its model information endpoint returned HTTP $1. Check that the service can report this model’s output limit. If it offers an OpenAI-compatible API, you can select that format and address in Model Settings.',
  ],
  [
    /^模型服务 (\S+) 拒绝连接（ECONNREFUSED）。请确认 API 已启动，且监听地址和端口与配置一致。$/,
    '$1 refused the model connection (ECONNREFUSED). Check that the API is running and listening at the configured address and port.',
  ],
  [
    /^无法解析模型服务 (\S+) 的地址。请检查主机名和 DNS 设置。$/,
    'The hostname $1 could not be resolved. Check the hostname and DNS settings.',
  ],
  [
    /^连接模型服务 (\S+) 超时。请检查 API 是否正常响应，以及网络或防火墙是否允许此连接。$/,
    'The connection to $1 timed out. Check that the API is responding and that the network and firewall allow the connection.',
  ],
  [
    /^模型服务 (\S+) 的连接被中断。请检查 API 日志，并确认地址使用的 HTTP 或 HTTPS 与服务一致。$/,
    'The connection to $1 was interrupted. Check the API logs and confirm that the URL uses the correct HTTP or HTTPS scheme.',
  ],
  [
    /^无法验证模型服务 (\S+) 的 HTTPS 证书。请检查证书有效期、主机名和 Node.js 的受信任证书配置。$/,
    'The HTTPS certificate for $1 could not be verified. Check its expiration date and hostname, and the trusted certificate settings for Node.js.',
  ],
  [
    /^模型服务 (\S+) 未返回有效的 HTTP 响应。请确认 HTTP 或 HTTPS、端口和 API 地址与模型服务一致。$/,
    '$1 did not return a valid HTTP response. Check the protocol, port and API address.',
  ],
  [
    /^本机服务无法访问 (\S+) 所在的网络。请检查网络连接、路由和防火墙。$/,
    'The local service cannot reach the network for $1. Check the connection, routes and firewall.',
  ],
  [
    /^本机服务无法连接模型服务 (\S+)。请检查 API 是否运行、服务地址及 API 日志。$/,
    'The local service could not connect to $1. Check that the API is running, verify its address and inspect the API logs.',
  ],
  [
    /^第 (\d+) 行需要 (\d+) 个单元格，请补齐空值或移除多余列。$/,
    'Row $1 must contain $2 cells. Fill in missing cells or remove extra columns.',
  ],
  [
    /^第 (\d+) 行第 (\d+) 列需要文本、数字、布尔值、null 或 \{"formula":"=公式"\}。$/,
    'Row $1, column $2 must contain text, a number, a boolean, null or a formula object such as {"formula":"=A1"}.',
  ],
  [
    /^第 (\d+) 行第 (\d+) 列包含公式，请将该列格式改为 general 或数字格式。$/,
    'Row $1, column $2 contains a formula. Set the column format to general or a numeric format.',
  ],
  [
    /^第 (\d+) 节表格第 (\d+) 行需要 (\d+) 个文本单元格，每格最多 500 字符。$/,
    'Section $1, table row $2 needs $3 text cells of up to 500 characters each.',
  ],
];

const fieldNames: Readonly<Record<string, string>> = {
  报告: 'Report',
  文档标题: 'Document title',
  数据表: 'Data table',
  表头: 'Column headers',
  工作表名称: 'Worksheet name',
  标题: 'title',
  段落: 'paragraphs',
  要点: 'bullets',
  表格: 'table',
};

const fieldProblems: readonly MessagePattern[] = [
  [/^需要 1–(\d+) 字符。$/, ' must contain 1–$1 characters.'],
  [/^最多 (\d+) 项。$/, ' can contain up to $1 items.'],
  [/^需要 1–(\d+) 个不重复的列名。$/, ' must contain 1–$1 unique column names.'],
  [/^格式无效。$/, ' has an invalid structure.'],
  [/^需要 1–60 行数据。$/, ' must contain 1–60 data rows.'],
  [/^需要正文、要点或表格。$/, ' requires at least one paragraph, bullet or table.'],
  [
    /^包含暂不支持的字段：([\s\S]+)。请将内容移入正文或表格后重试。$/,
    ' contains unsupported fields: $1. Move their content into the body or table, then try again.',
  ],
];

function matchPattern(message: string, patterns: readonly MessagePattern[]): string | undefined {
  for (const [pattern, english] of patterns) {
    if (pattern.test(message)) return message.replace(pattern, english);
  }
  return undefined;
}

/** Match only complete diagnostics emitted by the model and structured-output validators. */
export function translateModelMessage(message: string): string | undefined {
  const direct = Object.prototype.hasOwnProperty.call(modelUiMessages, message)
    ? modelUiMessages[message]
    : undefined;
  if (direct) return direct;

  const loopbackHint =
    ' 回环地址指运行 TrioBridge 本机服务的电脑；API 在容器或虚拟机中时，请确认端口已映射到这台电脑。';
  if (message.endsWith(loopbackHint)) {
    const reason = matchPattern(message.slice(0, -loopbackHint.length), variableMessages);
    if (reason)
      return (
        reason +
        ' Loopback points to the computer running TrioBridge. If the API runs in a container or VM, make sure its port is mapped to this computer.'
      );
  }

  // Shared settings can return both the credential-cleanup warning and the connection probe warning.
  const parts = message
    .split('。 ')
    .map((part, i, all) => (i < all.length - 1 ? part + '。' : part));
  if (
    parts.length > 1 &&
    parts.every((part) => Object.prototype.hasOwnProperty.call(modelUiMessages, part))
  )
    return parts.map((part) => modelUiMessages[part]).join(' ');

  const variable = matchPattern(message, variableMessages);
  if (variable) return variable;

  const cell = /^第 (\d+) 行第 (\d+) 列：(.+)$/u.exec(message);
  if (cell && Object.prototype.hasOwnProperty.call(modelUiMessages, cell[3]))
    return `Row ${cell[1]}, column ${cell[2]}: ${modelUiMessages[cell[3]]}`;

  const field =
    /^(?:(第 (\d+) 节)(标题|段落|要点|表格|表头)?|(报告|文档标题|数据表|表头|工作表名称))(?:第 (\d+) 项)?([\s\S]+)$/u.exec(
      message,
    );
  if (!field) return undefined;
  const problem = matchPattern(field[6], fieldProblems);
  if (!problem) return undefined;
  const sectionField = field[3] === '表头' ? 'column headers' : fieldNames[field[3]];
  const label = field[2]
    ? `Section ${field[2]}${sectionField ? ' ' + sectionField : ''}`
    : fieldNames[field[4]];
  return `${label}${field[5] ? ', item ' + field[5] : ''}${problem}`;
}
