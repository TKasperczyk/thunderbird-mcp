#!/usr/bin/env node

const { spawn } = require("node:child_process");
const path = require("node:path");

const bridgePath = path.resolve(__dirname, "..", "mcp-bridge.cjs");
let child = null;
let nextId = 1;
let buffered = "";
const pending = new Map();

function ensureBridge() {
  if (child) return child;
  child = spawn(process.execPath, [bridgePath], {
    cwd: path.dirname(bridgePath),
    stdio: ["pipe", "pipe", "inherit"],
  });
  child.stdout.setEncoding("utf8");
  child.stdout.on("data", chunk => {
    buffered += chunk;
    while (buffered.includes("\n")) {
      const newline = buffered.indexOf("\n");
      const line = buffered.slice(0, newline).trim();
      buffered = buffered.slice(newline + 1);
      if (!line) continue;
      const message = JSON.parse(line);
      if (message.id && pending.has(message.id)) {
        const { resolve, reject } = pending.get(message.id);
        pending.delete(message.id);
        if (message.error) reject(new Error(message.error.message));
        else resolve(message.result);
      }
    }
  });
  return child;
}

function request(method, params = {}) {
  const bridge = ensureBridge();
  const id = nextId++;
  return new Promise((resolve, reject) => {
    pending.set(id, { resolve, reject });
    bridge.stdin.write(`${JSON.stringify({ jsonrpc: "2.0", id, method, params })}\n`);
  });
}

function notification(method, params = {}) {
  ensureBridge().stdin.write(`${JSON.stringify({ jsonrpc: "2.0", method, params })}\n`);
}

async function tool(name, args) {
  const response = await request("tools/call", { name, arguments: args });
  const text = response?.content?.find(item => item.type === "text")?.text;
  if (!text) throw new Error(`${name} boş yanıt verdi`);
  const parsed = JSON.parse(text);
  if (parsed?.error) throw new Error(`${name}: ${parsed.error}`);
  return parsed;
}

function normalize(value) {
  return String(value || "")
    .toLocaleLowerCase("tr-TR")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/ı/g, "i");
}

const categoryRules = [
  {
    name: "Aksiyon",
    key: "$label1",
    pattern: /\b(action required|aksiyon gerekli|islem gerekli|yanit gerekli|response required|verify|verification|dogrula(?:yin|yiniz|ma)?|onayla(?:yin|yiniz|ma)?|confirm|urgent|acil|security alert|guvenlik uyarisi|suspicious|supheli|unusual sign[- ]?in|failed|failure|basarisiz|declined|reddedildi|overdue|past due|son odeme|due date|expires?|suresi dol|renewal required|attention required|imza gerekli|signature required)\b/,
  },
  {
    name: "Bankalar",
    key: "$label2",
    pattern: /\b(enpara|qnb|yapi kredi|yapikredi|garanti|garantibbva|is bankasi|isbank|akbank|ziraat|vakifbank|halkbank|denizbank|teb|kuveyt turk|kuveytturk|albaraka|fibabanka|papara|param|iyzico|paytr|bankasi|bankacilik|kredi karti|hesap ozeti|para transferi|eft|havale)\b|@(enpara|qnb|yapikredi|garanti|garantibbva|isbank|akbank|ziraatbank|vakifbank|halkbank|denizbank|teb)\./,
  },
  {
    name: "Ödemeler",
    key: "$label3",
    pattern: /\b(odeme|payment|paid|receipt|makbuz|tahsilat|transaction|islem bildirimi|charged|debit(?:ed)?|credit(?:ed)?|refund|iade|subscription|abonelik|renewed|yenilendi|balance|bakiye|pos|checkout)\b/,
  },
  {
    name: "Faturalar",
    key: "$label4",
    pattern: /\b(fatura|invoice|e[- ]?fatura|e[- ]?arsiv|bill(?:ing)?|statement|ekstre|hesap ozeti|varlik dokumu|abonelik bedeli|ucret bildirimi)\b/,
  },
  {
    name: "Siparişler",
    key: "$label5",
    pattern: /\b(siparis(?:iniz|in|i|ler(?:iniz)?)?|order|kargo(?:ya|nuz|nuzun)?|cargo|shipment|shipping|teslimat|delivery|dispatched|gonderiniz|gonderi takip|tracking(?: number)?|paketiniz|sepetiniz|purchase)\b/,
  },
  {
    name: "Randevular",
    key: "randevular",
    pattern: /\b(randevu(?:nuz|nuzu|nuzun|lar)?|appointment|meeting|toplanti|calendar|takvim|invitation|invite|davet|reservation|rezervasyon|booking|webinar|etkinlik|event reminder|gorusme|mülakat|mulakat)\b/,
  },
  {
    name: "Bültenler",
    key: "bultenler",
    pattern: /\b(newsletter|bulten|haberci|digest|weekly|haftalik|monthly|aylik|kampanya|campaign|promosyon|promotion|indirim|discount|firsat|offer|sale|duyuru|announcement|yenilikler|what'?s new|unsubscribe|abonelikten cik|trendler|ozetiniz)\b/,
  },
  {
    name: "Teknik",
    key: "teknik",
    pattern: /\b(github|gitlab|bitbucket|sentry|cloudflare|vercel|netlify|hetzner|digitalocean|aws|azure|google cloud|server|sunucu|hosting|domain|alan adi|dns|ssl|certificate|sertifika|backup|yedekleme|deploy|deployment|build|pipeline|ci\/cd|workflow|repository|repo|pull request|merge request|issue|api|webhook|uptime|downtime|incident|vulnerability|guvenlik acigi|database|veritabani|stripe radar|developer)\b/,
  },
];

const managedCategoryKeys = new Set(categoryRules.map(rule => rule.key));
const actionRule = categoryRules.find(rule => rule.name === "Aksiyon");
const primaryRuleOrder = [
  "Bankalar",
  "Faturalar",
  "Siparişler",
  "Randevular",
  "Ödemeler",
  "Bültenler",
  "Teknik",
].map(name => categoryRules.find(rule => rule.name === name));

const excludedFolder = /spam|junk|trash|cop kutusu|çöp kutusu|sent|gonderilen|gönderilen|draft|taslak|sablon|şablon|unsent/i;

function classify(message, folderName) {
  const haystack = normalize([
    message.subject,
    message.author,
    message.recipients,
    message.ccList,
    message.preview,
  ].join(" "));

  const matches = new Set();
  if (actionRule.pattern.test(haystack)) matches.add(actionRule.key);

  const primaryRule = primaryRuleOrder.find(category =>
    category.pattern.test(haystack) ||
    (category.name === "Bültenler" && /newsletter/i.test(folderName))
  );
  if (primaryRule) matches.add(primaryRule.key);
  return matches;
}

function chunks(items, size) {
  const result = [];
  for (let index = 0; index < items.length; index += size) {
    result.push(items.slice(index, index + size));
  }
  return result;
}

async function main() {
  const applyChanges = process.argv.includes("--apply");
  await request("initialize", {
    protocolVersion: "2025-06-18",
    capabilities: {},
    clientInfo: { name: "thunderbird-mail-organizer", version: "1.0" },
  });
  notification("notifications/initialized");

  const folderTable = await tool("listFolders", { format: "table" });
  const folders = folderTable.rows
    .map(row => Object.fromEntries(folderTable.columns.map((name, index) => [name, row[index]])))
    .filter(folder => folder.totalMessages > 0)
    .filter(folder => ["inbox", "folder"].includes(folder.type))
    .filter(folder => !excludedFolder.test(normalize(folder.name)));

  const changeGroups = new Map();
  const discoveredCounts = Object.fromEntries(categoryRules.map(rule => [rule.key, 0]));
  const proposedAddCounts = Object.fromEntries(categoryRules.map(rule => [rule.key, 0]));
  const proposedRemoveCounts = Object.fromEntries(categoryRules.map(rule => [rule.key, 0]));
  const pendingExamples = [];
  let scanned = 0;
  let matched = 0;

  for (const folder of folders) {
    let offset = 0;
    while (true) {
      const page = await tool("searchMessages", {
        query: "",
        folderPath: folder.path,
        includeSubfolders: false,
        dedupByMessageId: false,
        maxResults: 200,
        offset,
        sortOrder: "desc",
      });

      const messages = page.messages || [];
      scanned += messages.length;
      for (const message of messages) {
        const desiredKeys = classify(message, folder.name);
        if (desiredKeys.size > 0) matched += 1;
        for (const key of desiredKeys) discoveredCounts[key] += 1;

        const currentKeys = new Set(
          (message.tags || []).filter(key => managedCategoryKeys.has(key))
        );
        const addTags = [...desiredKeys]
          .filter(key => !currentKeys.has(key))
          .sort();
        const removeTags = [...currentKeys]
          .filter(key => !desiredKeys.has(key))
          .sort();
        if (addTags.length === 0 && removeTags.length === 0) continue;

        for (const key of addTags) proposedAddCounts[key] += 1;
        for (const key of removeTags) proposedRemoveCounts[key] += 1;
        if (pendingExamples.length < 20) {
          pendingExamples.push({
            messageId: message.id,
            messageKey: message.messageKey,
            subject: message.subject,
            folder: folder.name,
            folderPath: folder.path,
            addTags,
            removeTags,
          });
        }
        const targetField = Number.isInteger(message.messageKey)
          ? "messageKeys"
          : "messageIds";
        const target = targetField === "messageKeys" ? message.messageKey : message.id;
        const groupKey = JSON.stringify([folder.path, addTags, removeTags, targetField]);
        if (!changeGroups.has(groupKey)) changeGroups.set(groupKey, new Set());
        changeGroups.get(groupKey).add(target);
      }

      if (!page.hasMore || messages.length === 0) break;
      offset += messages.length;
    }
  }

  const appliedCounts = Object.fromEntries(categoryRules.map(rule => [rule.key, 0]));
  const removedCounts = Object.fromEntries(categoryRules.map(rule => [rule.key, 0]));
  if (applyChanges) {
    for (const [groupKey, targets] of changeGroups) {
      const [folderPath, addTags, removeTags, targetField] = JSON.parse(groupKey);
      for (const batch of chunks([...targets], 200)) {
        await tool("updateMessage", {
          [targetField]: batch,
          folderPath,
          addTags,
          removeTags,
        });
        for (const key of addTags) appliedCounts[key] += batch.length;
        for (const key of removeTags) removedCounts[key] += batch.length;
      }
    }
  }

  const verification = {};
  for (const category of categoryRules) {
    let total = 0;
    for (const folder of folders) {
      const count = await tool("searchMessages", {
        query: "",
        folderPath: folder.path,
        includeSubfolders: false,
        dedupByMessageId: false,
        tag: category.key,
        countOnly: true,
      });
      total += typeof count === "number" ? count : (count.count || count.totalMatches || 0);
    }
    verification[category.name] = {
      desired: discoveredCounts[category.key],
      proposedAdd: proposedAddCounts[category.key],
      proposedRemove: proposedRemoveCounts[category.key],
      applied: appliedCounts[category.key],
      removed: removedCounts[category.key],
      currentTotal: total,
    };
  }

  process.stdout.write(`${JSON.stringify({
    mode: applyChanges ? "apply" : "dry-run",
    note: applyChanges
      ? "Etiketler uygulandı ve yönetilen eski çakışmalar kaldırıldı."
      : "Hiçbir ileti değiştirilmedi. Uygulamak için --apply kullanın.",
    folders: folders.length,
    scanned,
    matched,
    unmatched: scanned - matched,
    changeGroups: changeGroups.size,
    pendingExamples,
    categories: verification,
  }, null, 2)}\n`);
}

if (require.main === module) {
  main()
    .catch(error => {
      process.stderr.write(`${error.stack || error.message}\n`);
      process.exitCode = 1;
    })
    .finally(() => child?.kill());
} else {
  module.exports = { categoryRules, classify, managedCategoryKeys, primaryRuleOrder };
}
