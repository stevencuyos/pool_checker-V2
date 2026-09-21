function splitPools(poolString) {
  if (!poolString) return [];
  const pools = poolString.split(',').map(p => p.trim()).filter(Boolean);
  return [...new Set(pools)];
}

function parseActivityTime(timeStr) {
  if (!timeStr) return null;
  // Match strictly format "yyyy-MM-dd HH:mm:ss"
  const regex = /^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})$/;
  const match = timeStr.match(regex);
  if (!match) {
    throw new Error(`Unparseable activity time: ${timeStr}`);
  }
  // Construct ISO string with +08:00
  const isoStr = `${match[1]}-${match[2]}-${match[3]}T${match[4]}:${match[5]}:${match[6]}+08:00`;
  return isoStr;
}

function parseExtractedAt(filename, fallbackDateObj) {
  const regex = /^Agent Report (\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})/;
  const match = filename.match(regex);
  if (match) {
    return `${match[1]}-${match[2]}-${match[3]}T${match[4]}:${match[5]}:${match[6]}+08:00`;
  }

  if (fallbackDateObj instanceof Date) {
    // Assuming fallbackDateObj is already in local timezone (Asia/Manila)
    // Or we format it explicitly
    // Date.toISOString() gives UTC. We want to output in +08:00 representation.
    // However, JS Date might be executing in UTC environment.
    // Let's format manually based on UTC+8 offset
    const timeMs = fallbackDateObj.getTime();
    const manilaMs = timeMs + 8 * 60 * 60 * 1000;
    const mDate = new Date(manilaMs);
    const pad = (n) => String(n).padStart(2, '0');
    return `${mDate.getUTCFullYear()}-${pad(mDate.getUTCMonth() + 1)}-${pad(mDate.getUTCDate())}T${pad(mDate.getUTCHours())}:${pad(mDate.getUTCMinutes())}:${pad(mDate.getUTCSeconds())}+08:00`;
  }
  return null;
}

function computeStaleness(dataEndIso, nowMs) {
  if (!dataEndIso) return { staleDays: 0, isStale: false };

  const endMs = new Date(dataEndIso).getTime();
  const manilaEndMs = endMs + 8 * 60 * 60 * 1000;
  const manilaNowMs = nowMs + 8 * 60 * 60 * 1000;

  const endDate = new Date(manilaEndMs);
  const nowDate = new Date(manilaNowMs);

  // Calculate calendar days in Asia/Manila (ignoring time)
  // By converting to UTC start of day
  const endDayStart = Date.UTC(endDate.getUTCFullYear(), endDate.getUTCMonth(), endDate.getUTCDate());
  const nowDayStart = Date.UTC(nowDate.getUTCFullYear(), nowDate.getUTCMonth(), nowDate.getUTCDate());

  const staleDays = Math.floor((nowDayStart - endDayStart) / (24 * 60 * 60 * 1000));

  return {
    staleDays: staleDays,
    isStale: staleDays > 1
  };
}

function dedupeLatest(rows, headers) {
  // Find column indices
  const agentIdx = headers.findIndex(h => h === "Agent");
  const managerIdx = headers.findIndex(h => h === "Manager");
  const locationIdx = headers.findIndex(h => h === "Location");

  const startIdx = headers.findIndex(h => h.startsWith("Activity Start Time"));
  const endIdx = headers.findIndex(h => h.startsWith("Activity End Time"));

  const phoneStatusIdx = headers.findIndex(h => h === "Phone Channel Status");
  const chatStatusIdx = headers.findIndex(h => h === "Chat Channel Status");
  const emailStatusIdx = headers.findIndex(h => h === "Email Channel Status");

  const phonePrimaryIdx = headers.findIndex(h => h === "Phone Primary Pools");
  const phoneSecondaryIdx = headers.findIndex(h => h === "Phone Secondary Pools");
  const chatPrimaryIdx = headers.findIndex(h => h === "Chat Primary Pools");
  const chatSecondaryIdx = headers.findIndex(h => h === "Chat Secondary Pools");
  const emailPrimaryIdx = headers.findIndex(h => h === "Email Primary Pools");
  const emailSecondaryIdx = headers.findIndex(h => h === "Email Secondary Pools");

  const agentsMap = new Map();

  let globalDataStartMs = Infinity;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const agent = row[agentIdx];
    if (!agent) continue;

    // Parse time
    let startTimeStr = row[startIdx];
    let endTimeStr = row[endIdx];
    let parsedEndTime = null;
    let endTimeMs = 0;
    try {
      if (startTimeStr) {
        let parsedStart = parseActivityTime(startTimeStr);
        let startMs = new Date(parsedStart).getTime();
        if (startMs < globalDataStartMs) {
          globalDataStartMs = startMs;
        }
      }
      parsedEndTime = parseActivityTime(endTimeStr);
      endTimeMs = new Date(parsedEndTime).getTime();
    } catch (e) {
      throw new Error(`Row ${i + 2}: ${e.message}`); // +2 assuming row 1 is headers
    }

    const poolStrings = [
      row[phonePrimaryIdx] || '',
      row[phoneSecondaryIdx] || '',
      row[chatPrimaryIdx] || '',
      row[chatSecondaryIdx] || '',
      row[emailPrimaryIdx] || '',
      row[emailSecondaryIdx] || ''
    ];

    const isOnline = (status) => status && status !== 'STATUS_OFFLINE';

    if (!agentsMap.has(agent)) {
      agentsMap.set(agent, {
        agent: agent,
        manager: row[managerIdx] || '',
        location: row[locationIdx] || '',
        latestEndMs: endTimeMs,
        latestEndIso: parsedEndTime,
        rowsInReport: 1,
        changedDuringDay: false,
        phoneSeenOnline: isOnline(row[phoneStatusIdx]),
        chatSeenOnline: isOnline(row[chatStatusIdx]),
        emailSeenOnline: isOnline(row[emailStatusIdx]),
        latestPoolStrings: poolStrings,
        latestRow: row
      });
    } else {
      const existing = agentsMap.get(agent);
      existing.rowsInReport++;

      existing.phoneSeenOnline = existing.phoneSeenOnline || isOnline(row[phoneStatusIdx]);
      existing.chatSeenOnline = existing.chatSeenOnline || isOnline(row[chatStatusIdx]);
      existing.emailSeenOnline = existing.emailSeenOnline || isOnline(row[emailStatusIdx]);

      // Check for pool changes
      for (let j = 0; j < 6; j++) {
        const pool1 = splitPools(existing.latestPoolStrings[j]).sort().join(',');
        const pool2 = splitPools(poolStrings[j]).sort().join(',');
        if (pool1 !== pool2) {
          existing.changedDuringDay = true;
        }
      }

      if (endTimeMs > existing.latestEndMs) {
        existing.latestEndMs = endTimeMs;
        existing.latestEndIso = parsedEndTime;

        // update latest row and strings, but preserve changedDuringDay if it was already true
        existing.latestPoolStrings = poolStrings;
        existing.latestRow = row;
        existing.manager = row[managerIdx] || '';
        existing.location = row[locationIdx] || '';
      }
    }
  }

  return {
    agents: Array.from(agentsMap.values()),
    globalDataStartMs
  };
}

if (typeof module !== 'undefined') {
  module.exports = {
    splitPools,
    parseActivityTime,
    parseExtractedAt,
    computeStaleness,
    dedupeLatest
  };
}
