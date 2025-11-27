var scriptProperties = PropertiesService.getScriptProperties();

var OPENROUTER_API_KEY = scriptProperties.getProperty('OPENROUTER_API_KEY');
var OPENAI_API_KEY = scriptProperties.getProperty('OPENAI_API_KEY');
var GEMINI_API_KEY = scriptProperties.getProperty('GEMINI_API_KEY');
var DEEPSEEK_API_KEY = scriptProperties.getProperty('DEEPSEEK_API_KEY');

const OPENROUTER_POOL = [
  { model: "openai/gpt-4o-mini", supportsJson: true },
  { model: "anthropic/claude-3-haiku", supportsJson: true },
  { model: "meta-llama/llama-3.1-70b-instruct", supportsJson: true },
  { model: "google/gemini-flash-1.5", supportsJson: true }
];

const DIRECT_PROVIDERS = [
  {
    name: "Backup:OpenAI",
    enabled: !!OPENAI_API_KEY,
    url: "https://api.openai.com/v1/chat/completions",
    headers: {
      Authorization: "Bearer " + OPENAI_API_KEY,
      "Content-Type": "application/json"
    },
    buildPayload: function (prompt) {
      return {
        model: "gpt-4o-mini",
        messages: [
          { role: "system", content: "You are a helpful assistant that outputs only JSON." },
          { role: "user", content: prompt }
        ],
        temperature: 0
      };
    },
    parseResponse: function (json) {
      return json.choices[0].message.content;
    }
  },
  {
    name: "Backup:Gemini",
    enabled: !!GEMINI_API_KEY,
    url: "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + GEMINI_API_KEY,
    headers: { "Content-Type": "application/json" },
    buildPayload: function (prompt) {
      return {
        contents: [{ parts: [{ text: prompt }] }]
      };
    },
    parseResponse: function (json) {
      return json.candidates[0].content.parts[0].text;
    }
  },
  {
    name: "Backup:DeepSeek",
    enabled: !!DEEPSEEK_API_KEY,
    url: "https://api.deepseek.com/chat/completions",
    headers: {
      Authorization: "Bearer " + DEEPSEEK_API_KEY,
      "Content-Type": "application/json"
    },
    buildPayload: function (prompt) {
      return {
        model: "deepseek-chat",
        messages: [
          { role: "system", content: "You are a helpful assistant that outputs only JSON." },
          { role: "user", content: prompt }
        ],
        temperature: 0
      };
    },
    parseResponse: function (json) {
      return json.choices[0].message.content;
    }
  }
];

function explainLeetCodeSelection() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var selection = sheet.getSelection().getActiveRange();

  if (!selection) {
    SpreadsheetApp.getUi().alert("請先選擇包含 LeetCode 題目名稱的儲存格。");
    return;
  }

  var startRow = selection.getRow();
  var numRows = selection.getNumRows();
  var values = selection.getValues();

  ensureHeaders(sheet);

  var results = new Array(numRows).fill(null).map(function () {
    return { success: false, data: null, provider: "", error: "" };
  });

  ss.toast("開始分析 " + numRows + " 個題目...", "LeetCode AI", 20);

  if (OPENROUTER_API_KEY) {
    try {
      fetchViaOpenRouterBatch(values, results);
    } catch (err) {
      console.error("OpenRouter Critical Failure: " + err.message);
    }
  } else {
    console.warn("No OpenRouter Key found, skipping to backup.");
  }

  var failedCount = results.filter(function (res) {
    return !res.success && res.error !== "Empty Title";
  }).length;

  if (failedCount > 0) {
    ss.toast("OpenRouter 部分失敗，正在切換備援線路 (" + failedCount + " 題)...", "System Alert", 5);
    fetchViaDirectBackup(values, results);
  }

  var providerValues = [];
  var detailValues = [];

  for (var i = 0; i < numRows; i++) {
    var res = results[i];
    var title = values[i][0];

    if (!title || title.toString().trim() === "") {
      providerValues.push([""]);
      detailValues.push(["", "", "", "", "", "", ""]);
    } else if (res.success) {
      providerValues.push([res.provider]);
      detailValues.push([
        res.data.G_Problem_Explain,
        res.data.H_App_Scenario_Algo,
        res.data.I_Title_ZH,
        res.data.J_Explain_ZH,
        res.data.K_Scenario_Technique_ZH,
        res.data.L_Time_Complexity,
        res.data.M_Space_Complexity
      ]);
    } else {
      providerValues.push(["Failed"]);
      detailValues.push([
        "Error: " + res.error,
        "",
        "",
        "",
        "",
        "",
        ""
      ]);
    }
  }

  sheet.getRange(startRow, 9, numRows, 1).setValues(providerValues);
  sheet.getRange(startRow, 10, numRows, 7).setValues(detailValues);

  SpreadsheetApp.flush();
  ss.toast("分析完成！", "完成", 3);
}

function fetchViaOpenRouterBatch(values, results) {
  var requests = [];
  var indices = [];
  var modelsUsed = [];

  for (var i = 0; i < values.length; i++) {
    var title = values[i][0];
    if (!title || title.toString().trim() === "") {
      results[i].error = "Empty Title";
      continue;
    }

    var poolEntry = OPENROUTER_POOL[Math.floor(Math.random() * OPENROUTER_POOL.length)];
    var prompt = generatePrompt(title);

    var payload = {
      model: poolEntry.model,
      messages: [
        { role: "system", content: "You are a helpful assistant that outputs only JSON." },
        { role: "user", content: prompt }
      ],
      temperature: 0
    };

    if (poolEntry.supportsJson) {
      payload.response_format = { type: "json_object" };
    }

    requests.push({
      url: "https://openrouter.ai/api/v1/chat/completions",
      method: "post",
      headers: {
        Authorization: "Bearer " + OPENROUTER_API_KEY,
        "Content-Type": "application/json",
        "HTTP-Referer": "https://script.google.com/",
        "X-Title": "Google Sheets LeetCode"
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    indices.push(i);
    modelsUsed.push(poolEntry.model);
  }

  if (!requests.length) {
    return;
  }

  var responses = UrlFetchApp.fetchAll(requests);

  for (var j = 0; j < responses.length; j++) {
    var idx = indices[j];
    var response = responses[j];
    var code = response.getResponseCode();

    if (code === 200) {
      try {
        var json = JSON.parse(response.getContentText());
        var content = json.choices[0].message.content;
        var parsedData = parseJsonResult(content);
        var usedModel = json.model || modelsUsed[j] || "OpenRouter";

        results[idx] = {
          success: true,
          data: parsedData,
          provider: usedModel,
          error: ""
        };
      } catch (err) {
        results[idx].error = "OpenRouter Parse Error";
      }
    } else {
      results[idx].error = "OpenRouter API Error " + code;
    }
  }
}

function fetchViaDirectBackup(values, results) {
  var activeProviders = DIRECT_PROVIDERS.filter(function (provider) {
    return provider.enabled;
  });

  if (!activeProviders.length) {
    console.warn("No backup providers enabled.");
    return;
  }

  var providerOrder = shuffle(activeProviders.slice());

  for (var p = 0; p < providerOrder.length; p++) {
    var provider = providerOrder[p];
    var pendingIndices = [];

    for (var i = 0; i < results.length; i++) {
      var title = values[i][0];
      if (!results[i].success && title && title.toString().trim() !== "") {
        pendingIndices.push(i);
      }
    }

    if (!pendingIndices.length) {
      break;
    }

    var requests = pendingIndices.map(function (index) {
      var prompt = generatePrompt(values[index][0]);
      return {
        index: index,
        request: {
          url: provider.url,
          method: "post",
          headers: provider.headers,
          payload: JSON.stringify(provider.buildPayload(prompt)),
          muteHttpExceptions: true
        }
      };
    });

    var responses = UrlFetchApp.fetchAll(
      requests.map(function (cfg) {
        return cfg.request;
      })
    );

    for (var r = 0; r < responses.length; r++) {
      var cfg = requests[r];
      var response = responses[r];
      var idx = cfg.index;

      if (response.getResponseCode() === 200) {
        try {
          var json = JSON.parse(response.getContentText());
          var content = provider.parseResponse(json);
          var parsedData = parseJsonResult(content);

          results[idx] = {
            success: true,
            data: parsedData,
            provider: provider.name,
            error: ""
          };
        } catch (err) {
          results[idx].error = provider.name + " Parse Error";
        }
      } else {
        results[idx].error = provider.name + " API Error " + response.getResponseCode();
      }
    }
  }

  for (var j = 0; j < results.length; j++) {
    if (!results[j].success && results[j].error === "") {
      results[j].error = "Backup providers failed";
    }
  }
}

function generatePrompt(title) {
  return (
    'You are an expert software engineer and algorithm instructor.\n' +
    'Analyze the LeetCode problem: "' + title + '".\n\n' +
    "Provide the output strictly in valid JSON format without Markdown code blocks.\n" +
    "The JSON must contain exactly these keys:\n" +
    '{\n' +
    '  "G_Problem_Explain": "Brief explanation of the problem in English.",\n' +
    '  "H_App_Scenario_Algo": "Real-world application scenario and the core algorithm/technique used (in English).",\n' +
    '  "I_Title_ZH": "The problem title translated to Traditional Chinese.",\n' +
    '  "J_Explain_ZH": "Brief explanation of the problem in Traditional Chinese.",\n' +
    '  "K_Scenario_Technique_ZH": "用淺顯易懂的方式解釋真實世界中這個題目的使用場景與時機並說明核心解題技巧 in Traditional Chinese.",\n' +
    '  "L_Time_Complexity": "說明最優解法的Time Complexity與其計算邏輯(e.g., O(n) 因為...).",\n' +
    '  "M_Space_Complexity": "說明最優解法的Space Complexity與其計算邏輯(e.g., O(n) 因為...)."\n' +
    "}\n"
  );
}

function parseJsonResult(text) {
  var cleanText = text.replace(/```json/g, "").replace(/```/g, "").trim();
  try {
    return JSON.parse(cleanText);
  } catch (err) {
    throw new Error("Invalid JSON format.");
  }
}

function shuffle(array) {
  for (var i = array.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }
  return array;
}

function ensureHeaders(sheet) {
  sheet.getRange(1, 9, 1, 8).setValues([[
    "LLM Provider",
    "Problem Explain",
    "Application Scenario & Core Algorithm/Technique",
    "題目名稱",
    "題目解釋",
    "應用場景+核心解題技巧",
    "Time Complexity",
    "Space Complexity"
  ]]);
}