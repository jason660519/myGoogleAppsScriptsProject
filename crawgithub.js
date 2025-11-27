var scriptProperties = PropertiesService.getScriptProperties();

// Helper function to safely get and trim properties
function getPropertySafe(key) {
  var val = scriptProperties.getProperty(key);
  return val ? val.toString().trim() : null;
}

var GITHUB_TOKEN = getPropertySafe("GITHUB_TOKEN");

// API Keys
var OPENROUTER_API_KEY = getPropertySafe("OPENROUTER_API_KEY");
var OPENAI_API_KEY = getPropertySafe("OPENAI_API_KEY");
var GEMINI_API_KEY = getPropertySafe("GEMINI_API_KEY");
var DEEPSEEK_API_KEY = getPropertySafe("DEEPSEEK_API_KEY");

var COL = { F: 6, G: 7, H: 8, I: 9, O: 15, P: 16, Q: 17 };

// --- 1. 定義 LLM Providers ---
const LLM_PROVIDERS = [
  {
    name: "Gemini",
    enabled: !!GEMINI_API_KEY,
    url: "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + GEMINI_API_KEY,
    headers: { "Content-Type": "application/json" },
    buildPayload: function (prompt) {
      return {
        contents: [{ parts: [{ text: prompt }] }],
        safetySettings: [
            { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" },
            { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_NONE" },
            { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" },
            { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_NONE" }
        ]
      };
    },
    parseResponse: function (json) {
      try {
        var parts = json && json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts;
        if (!parts || !parts.length) return null;
        return parts.map(p => p.text).join("").trim();
      } catch (e) { return null; }
    },
  },
  {
    name: "OpenRouter",
    enabled: !!OPENROUTER_API_KEY,
    url: "https://openrouter.ai/api/v1/chat/completions",
    headers: {
      Authorization: "Bearer " + OPENROUTER_API_KEY,
      "Content-Type": "application/json",
      "HTTP-Referer": "https://script.google.com/",
      "X-Title": "Google Sheets Github Tool",
    },
    buildPayload: function (prompt) {
      return {
        model: "openai/gpt-4o-mini",
        messages: [{ role: "system", content: "You are a helpful assistant that outputs only JSON." }, { role: "user", content: prompt }],
        response_format: { type: "json_object" },
      };
    },
    parseResponse: function (json) {
      try {
        var msg = json && json.choices && json.choices[0] && json.choices[0].message;
        if (!msg) return null;
        return msg.content || null;
      } catch (e) { return null; }
    },
  },
  {
    name: "DeepSeek",
    enabled: !!DEEPSEEK_API_KEY,
    url: "https://api.deepseek.com/chat/completions",
    headers: {
      Authorization: "Bearer " + DEEPSEEK_API_KEY,
      "Content-Type": "application/json",
    },
    buildPayload: function (prompt) {
      return {
        model: "deepseek-chat",
        messages: [{ role: "system", content: "You are a helpful assistant that outputs only JSON." }, { role: "user", content: prompt }],
        temperature: 0,
        response_format: { type: "json_object" },
      };
    },
    parseResponse: function (json) {
      try {
        var msg = json && json.choices && json.choices[0] && json.choices[0].message;
        if (!msg) return null;
        return msg.content || null;
      } catch (e) { return null; }
    },
  },
  {
    name: "OpenAI",
    enabled: !!OPENAI_API_KEY,
    url: "https://api.openai.com/v1/chat/completions",
    headers: {
      Authorization: "Bearer " + OPENAI_API_KEY,
      "Content-Type": "application/json",
    },
    buildPayload: function (prompt) {
      return {
        model: "gpt-4o-mini",
        messages: [{ role: "system", content: "You are a helpful assistant that outputs only JSON." }, { role: "user", content: prompt }],
        temperature: 0,
      };
    },
    parseResponse: function (json) {
      try {
        var msg = json && json.choices && json.choices[0] && json.choices[0].message;
        if (!msg) return null;
        return msg.content || null;
      } catch (e) { return null; }
    },
  },
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("GitHub")
    .addItem("抓取選取", "crawlGithubSelection")
    .addToUi();
}

function crawlGithubSelection() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var selection = sheet.getActiveRange();
  if (!selection) return;

  var startRow = selection.getRow();
  var values = selection.getValues();
  var cache = CacheService.getScriptCache();

  var activeProviders = LLM_PROVIDERS.filter(p => p.enabled);
  if (activeProviders.length === 0) {
    SpreadsheetApp.getUi().alert("未設定任何 LLM API Key");
    return;
  }

  var providerIndex = 0; 

  ss.toast("開始分析...", "GitHub Crawler", 10);

  for (var i = 0; i < values.length; i++) {
    var rowProcessed = false;

    try {
      for (var j = 0; j < values[i].length; j++) {
        var url = values[i][j];
        if (!url) continue;
        var s = url.toString().trim();
        if (s === "" || s.indexOf("github.com") === -1) continue;

        var row = startRow + i;
        var f = sheet.getRange(row, COL.F);
        var g = sheet.getRange(row, COL.G);
        var h = sheet.getRange(row, COL.H);
        var ii = sheet.getRange(row, COL.I);
        var opq = sheet.getRange(row, COL.O, 1, 3);

        // 1. 解析 URL
        var parsed = parseGithubUrl_Github(s);
        if (!parsed) {
          h.setValue("Invalid URL");
          continue;
        }

        // --- A. 抓取 GitHub Meta Data (API + HTML Scraping) ---
        // 升級 Cache Key 到 v10，強制刷新 O 欄位
        var cacheKey = "meta_v10:" + parsed.owner + "/" + parsed.repo;
        var cached = cache.get(cacheKey);
        var meta = null;
        
        if (cached) {
          try { meta = JSON.parse(cached); } catch (e) {}
        }
        
        if (!meta) {
          meta = fetchRepoMeta_Github(parsed.owner, parsed.repo, GITHUB_TOKEN);
          try { cache.put(cacheKey, JSON.stringify(meta), 21600); } catch (e) {}
        }
        
        updateBasicColumns_Github(f, g, opq, meta, s);

        // --- B. 準備 LLM 分析 (H & I 欄位) ---
        var curH = h.getValue().toString().trim();
        var curI = ii.getValue().toString().trim();

        var shouldRun = (curH === "" || curI === "" || curH === "-" || curI === "-" || curH.indexOf("Error") === 0 || curH.indexOf("Analyzing") === 0);

        if (shouldRun) {
          h.setValue("Analyzing...");
          SpreadsheetApp.flush();

          var contextText = "";
          var source = "API_RAW";

          var readmeRaw = fetchReadmeRaw_Github(parsed.owner, parsed.repo, GITHUB_TOKEN);

          if (readmeRaw && readmeRaw.length > 100) {
            contextText = readmeRaw;
          } else {
            source = "WEB";
            var webText = fetchRepoPageText_Github(s);
            if (webText && webText.length > 100) {
              contextText = webText;
            } else {
              source = "META_ONLY";
              contextText = "Project Description: " + (meta.about || "No description") + 
                            "\nTopics: " + (meta.topics || "No topics") +
                            "\nLanguage: " + (meta.languages || "Unknown");
            }
          }

          if (contextText && contextText.length > 5) {
            var prompt = generatePrompt_Github(parsed.repo, contextText);
            var success = false;
            var lastError = "";
            
            var attempts = 0;
            var currentProviderIndex = providerIndex;
            
            while (attempts < activeProviders.length) {
                var currentProvider = activeProviders[currentProviderIndex];
                
                ss.toast("分析中 (" + currentProvider.name + "): " + parsed.repo, "AI Processing", 3);

                var result = callLLM_Github(currentProvider, prompt);
                
                if (result.success) {
                    var finalData = normalizeKeys_Github(result.data);
                    
                    if (finalData.features) h.setValue(finalData.features);
                    else h.setValue("Error: Missing 'features' key.");

                    if (finalData.scenarios) ii.setValue(finalData.scenarios);
                    else ii.setValue("-");

                    success = true;
                    break; 
                } else {
                    lastError = result.error;
                    console.warn("Provider " + currentProvider.name + " failed: " + lastError);
                    currentProviderIndex = (currentProviderIndex + 1) % activeProviders.length;
                    attempts++;
                }
            }

            rowProcessed = true;

            if (!success) {
              h.setValue("Error: All Failed (" + lastError + ")");
              ii.setValue("Check Logs");
            }
            
          } else {
            h.setValue("Error: No Context Found");
            ii.setValue("-");
          }
        }
      }
    } catch (err) {
      console.error("Row Error: " + err);
      try {
         var row = startRow + i;
         sheet.getRange(row, COL.H).setValue("Script Error: " + err.message);
      } catch(e) {}
    }

    if (rowProcessed) {
      var sleepTime = Math.floor(Math.random() * 2000) + 1000;
      Utilities.sleep(sleepTime);
    } else {
      if ((i + 1) % 5 === 0) Utilities.sleep(500);
    }
  }
  SpreadsheetApp.getUi().alert("完成");
}

// --- 輔助函數 ---
function normalizeKeys_Github(obj) {
  var newObj = { features: "", scenarios: "" };
  if (!obj) return newObj;
  
  for (var key in obj) {
    var lowerKey = key.toLowerCase();
    var val = obj[key];
    
    if (lowerKey.includes("feature") || lowerKey.includes("function") || lowerKey.includes("summary") || lowerKey.includes("desc")) {
      if (!newObj.features && typeof val === 'string') newObj.features = val;
    } 
    else if (lowerKey.includes("scenario") || lowerKey.includes("usage") || lowerKey.includes("case")) {
      if (!newObj.scenarios && typeof val === 'string') newObj.scenarios = val;
    }
  }
  
  if (!newObj.features) {
      var keys = Object.keys(obj);
      if (keys.length > 0 && typeof obj[keys[0]] === 'string') {
          newObj.features = obj[keys[0]];
      }
  }
  
  return newObj;
}

function callLLM_Github(provider, prompt) {
  try {
    var payload = provider.buildPayload(prompt);
    var options = {
      method: "post",
      headers: provider.headers,
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };

    var response = UrlFetchApp.fetch(provider.url, options);
    var code = response.getResponseCode();

    if (code === 200) {
      var json = JSON.parse(response.getContentText());
      var rawContent = provider.parseResponse(json);
      if (!rawContent) return { success: false, error: provider.name + ": Null Response" };
      
      var parsedData = parseJsonResult_Github(rawContent);
      if (parsedData) {
        return { success: true, data: parsedData };
      } else {
        return { success: false, error: provider.name + " JSON Parse Fail" };
      }
    } else {
      return { success: false, error: provider.name + " Error " + code };
    }
  } catch (e) {
    return { success: false, error: provider.name + " Exception: " + e.message };
  }
}

function generatePrompt_Github(repoName, contentText) {
  var truncatedContent = contentText.substring(0, 12000);
  return (
    "You are a technical editor.\n" +
    'Analyze the GitHub repository: "' + repoName + '".\n' +
    "Context:\n" + truncatedContent + "\n\n" +
    "Provide the output strictly in valid JSON format.\n" +
    "The JSON must contain exactly these keys:\n" +
    "{\n" +
    '  "features": "該Github Repo的主要功能與特色(簡短繁中文介紹不超過100字)",\n' +
    '  "scenarios": "該Github Repo的使用場景說明(繁中文不超過600字)"\n' +
    "}\n"
  );
}

function parseJsonResult_Github(text) {
  if (!text) return null;
  if (typeof text === "object") return text;
  
  var s = String(text);
  try { return JSON.parse(s); } catch (e) {}

  var cleanText = s.replace(/```json/g, "").replace(/```/g, "").trim();
  try { return JSON.parse(cleanText); } catch (e) {}

  var firstOpen = s.indexOf("{");
  var lastClose = s.lastIndexOf("}");
  if (firstOpen !== -1 && lastClose !== -1 && lastClose > firstOpen) {
    var jsonString = s.substring(firstOpen, lastClose + 1);
    try { return JSON.parse(jsonString); } catch (e) {}
  }

  return null;
}

// --- GitHub 資料抓取函數 ---

function fetchReadmeRaw_Github(owner, repo, token) {
  var headers = { "User-Agent": "Google-Apps-Script", Accept: "application/vnd.github.v3.raw" };
  if (token) headers["Authorization"] = "token " + token;
  try {
    var url = "https://api.github.com/repos/" + owner + "/" + repo + "/readme";
    var res = UrlFetchApp.fetch(url, { method: "get", headers: headers, muteHttpExceptions: true });
    if (res.getResponseCode() === 200) return res.getContentText();
  } catch (e) {}
  return null;
}

function fetchRepoPageText_Github(url) {
  try {
    var headers = { "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36" };
    var res = UrlFetchApp.fetch(url, { headers: headers, muteHttpExceptions: true });
    if (res.getResponseCode() === 200) return stripHtml_Github(res.getContentText());
  } catch (e) {}
  return null;
}

function stripHtml_Github(html) {
  if (!html) return "";
  var text = html.replace(/<script[^>]*>([\S\s]*?)<\/script>/gim, "");
  text = text.replace(/<style[^>]*>([\S\s]*?)<\/style>/gim, "");
  text = text.replace(/<[^>]+>/g, " ");
  text = text.replace(/\s+/g, " ").trim();
  return text;
}

// --- 修正後的 Meta 抓取 (HTML Scraping 優先 + API Fallback) ---
function fetchRepoMeta_Github(owner, repo, token) {
  var headers = { "User-Agent": "Google-Apps-Script" };
  if (token) headers["Authorization"] = "token " + token;
  var options = { method: "get", headers: headers, muteHttpExceptions: true };

  // 1. 抓取基本資訊 (API)
  var repoUrl = "https://api.github.com/repos/" + owner + "/" + repo;
  var repoRes = UrlFetchApp.fetch(repoUrl, options);

  if (repoRes.getResponseCode() !== 200) {
    return {
      homepage: "-",
      about: "Error " + repoRes.getResponseCode(),
      contributors: "-",
      license: "-",
      languages: "-",
      topics: "",
    };
  }

  var repoJson = JSON.parse(repoRes.getContentText());
  var about = repoJson.description || "";
  var homepage = repoJson.homepage || "";
  var topics = repoJson.topics ? repoJson.topics.join(", ") : "";
  var license = repoJson.license && repoJson.license.name ? repoJson.license.name : "No License";

  // 2. 抓取 Contributors (優先使用 HTML Scraping 以獲得網頁版精確數字)
  var contributors = "Check Repo";
  var scraped = false;

  try {
    // 嘗試爬取網頁 HTML
    var htmlUrl = "https://github.com/" + owner + "/" + repo;
    var htmlHeaders = {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    };
    var htmlRes = UrlFetchApp.fetch(htmlUrl, { headers: htmlHeaders, muteHttpExceptions: true });
    
    if (htmlRes.getResponseCode() === 200) {
      var html = htmlRes.getContentText();
      // Regex: 尋找 href 包含 contributors 的連結，並抓取其後的 Counter span 內容
      // 範例: <a href="/owner/repo/graphs/contributors" ...> ... <span class="Counter ...">117</span>
      var regex = /href=["']\/[^\/]+\/[^\/]+\/graphs\/contributors["'][^>]*>[\s\S]*?class=["'][^"']*Counter[^"']*["'][^>]*>([\d,]+)<\/span>/i;
      var match = html.match(regex);
      
      if (match && match[1]) {
        contributors = match[1].replace(/,/g, ""); // 移除千分位逗號
        scraped = true;
      } else {
        // 備用 Regex: 有時候連結沒有 graphs
        var regex2 = /href=["']\/[^\/]+\/[^\/]+\/contributors["'][^>]*>[\s\S]*?class=["'][^"']*Counter[^"']*["'][^>]*>([\d,]+)<\/span>/i;
        var match2 = html.match(regex2);
        if (match2 && match2[1]) {
           contributors = match2[1].replace(/,/g, "");
           scraped = true;
        }
      }
    }
  } catch (e) {
    // Scraping 失敗，靜默處理，交給下方的 API Fallback
  }

  // 如果爬蟲失敗，回退到 API Pagination 方法
  if (!scraped) {
    try {
      Utilities.sleep(100); 
      var contribUrl = "https://api.github.com/repos/" + owner + "/" + repo + "/contributors?per_page=1&anon=true";
      var contribRes = UrlFetchApp.fetch(contribUrl, options);
      
      if (contribRes.getResponseCode() === 200) {
        var headers = contribRes.getHeaders();
        var linkHeader = headers['Link'] || headers['link'];
        
        if (linkHeader) {
          var match = linkHeader.match(/[?&]page=(\d+)[^>]*>;\s*rel="last"/);
          if (match) {
            contributors = match[1];
          } else {
            var json = JSON.parse(contribRes.getContentText());
            contributors = json.length.toString();
          }
        } else {
          var json = JSON.parse(contribRes.getContentText());
          contributors = json.length.toString();
        }
      }
    } catch (e) {
      contributors = "Error"; 
    }
  }

  // 3. 抓取詳細語言資訊
  var languages = "Mixed";
  try {
    Utilities.sleep(100);
    var langUrl = "https://api.github.com/repos/" + owner + "/" + repo + "/languages";
    var langRes = UrlFetchApp.fetch(langUrl, options);
    if (langRes.getResponseCode() === 200) {
      var langJson = JSON.parse(langRes.getContentText());
      var totalBytes = 0;
      for (var key in langJson) {
        totalBytes += langJson[key];
      }
      if (totalBytes > 0) {
        var langArr = [];
        for (var key in langJson) {
          var pct = ((langJson[key] / totalBytes) * 100).toFixed(1);
          if (pct > 0.0) {
             langArr.push(key + " " + pct + "%");
          }
        }
        languages = langArr.join(", ");
      } else {
        languages = repoJson.language || "None";
      }
    } else {
      languages = repoJson.language || "Unknown";
    }
  } catch (e) {
    languages = repoJson.language || "Error";
  }

  return {
    homepage: homepage,
    about: about,
    contributors: contributors,
    license: license,
    languages: languages,
    topics: topics,
  };
}

function updateBasicColumns_Github(f, g, opq, meta, originalUrl) {
  var curF = f.getRichTextValue();
  var curFText = curF ? curF.getText() : "";
  var curFLink = curF ? curF.getLinkUrl() : null;
  var curG = g.getValue();
  var curOPQ = opq.getValues()[0];

  if (!curFLink) {
    if (curFText !== "" && meta.homepage && meta.homepage.indexOf("http") === 0) {
      var rt = SpreadsheetApp.newRichTextValue().setText(curFText).setLinkUrl(meta.homepage).build();
      f.setRichTextValue(rt);
    } else if (curFText === "" && meta.homepage && meta.homepage.indexOf("http") === 0) {
      f.setValue(meta.homepage);
    } else if (curFText === "") {
      f.setValue(originalUrl);
    }
  }
  
  if ((curG === null || curG.toString().trim() === "") && meta.about) {
    g.setValue(meta.about);
  }

  var outContrib = curOPQ[0] || meta.contributors;
  var outLic = curOPQ[1] || meta.license;
  var outLang = curOPQ[2] || meta.languages;

  opq.setValues([[outContrib, outLic, outLang]]);
}

function parseGithubUrl_Github(url) {
  var u = url.trim();
  if (u.indexOf("http") !== 0) u = "https://" + u;
  u = u.replace(/\.git$/, "").replace(/\/$/, "");
  try {
    var parts = u.split("/");
    var ghIndex = -1;
    for(var i=0; i<parts.length; i++) {
      if(parts[i].indexOf("github.com") !== -1) {
        ghIndex = i;
        break;
      }
    }
    if (ghIndex !== -1 && parts.length >= ghIndex + 3) {
      return { owner: parts[ghIndex + 1], repo: parts[ghIndex + 2] };
    }
  } catch(e) {}
  return null;
}