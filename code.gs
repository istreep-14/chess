
/**
 * Chess.com Complete Game Data Fetcher using Public API
 * Collects ALL possible game data fields with Me/Opponent formatting
 * 
 * Setup Instructions:
 * 1. Open Google Apps Script (script.google.com)
 * 2. Create a new project and paste this code
 * 3. Replace 'YOUR_USERNAME' with your Chess.com username
 * 4. Save the project
 * 5. Create a Google Sheet and note the Sheet ID from the URL
 * 6. Replace 'YOUR_SHEET_ID' with your actual Sheet ID
 * 7. Run the function or set up a trigger
 */

// Configuration - UPDATE THESE VALUES
const CHESS_COM_USERNAME = 'ians141'; // Replace with your Chess.com username
const SHEET_ID = '1OAQz_2Ev2lYMoiNiDQnXhicOKHtSu7xqZK2PEudqN2I'; // Replace with your Google Sheet ID
const SHEET_NAME = 'Games'; // Name of the sheet tab
const MOVES_SHEET_NAME = 'Moves'; // Name of the moves detail sheet
const MOVES_SAN_SHEET_NAME = 'Moves SAN';
const MOVES_CLOCK_SHEET_NAME = 'Moves Clock';

/**
 * Main function to fetch all Chess.com games with complete data
 */
function fetchAllChessComGames() {
  try {
    console.log('Starting to fetch Chess.com games with complete data...');
    
    // Get or create the sheet
    const sheet = getOrCreateSheet();
    
    // Clear existing data (optional - remove if you want to append)
    sheet.clear();
    
    // Set up headers
    setupHeaders(sheet);
    
    // Get all game archives (monthly archives)
    const archives = getGameArchives();
    console.log(`Found ${archives.length} monthly archives`);
    
    let totalGames = 0;
    let allGamesData = [];
    
    // Fetch games from each archive
    for (let i = 0; i < archives.length; i++) {
      const archiveUrl = archives[i];
      console.log(`Fetching archive ${i + 1}/${archives.length}: ${archiveUrl}`);
      
      const gamesInArchive = fetchGamesFromArchive(archiveUrl);
      allGamesData = allGamesData.concat(gamesInArchive);
      totalGames += gamesInArchive.length;
      
      // Add a small delay to avoid rate limiting
      if (i < archives.length - 1) {
        Utilities.sleep(500);
      }
    }
    
    // Order newest -> oldest by endTime then Game ID
    allGamesData.sort((a, b) => {
      const ta = a.endTime instanceof Date ? a.endTime.getTime() : 0;
      const tb = b.endTime instanceof Date ? b.endTime.getTime() : 0;
      if (tb !== ta) return tb - ta;
      const ai = Number(a.gameId); const bi = Number(b.gameId);
      if (Number.isFinite(ai) && Number.isFinite(bi)) return bi - ai;
      return 0;
    });
    
    // Write all games data to sheet
    if (allGamesData.length > 0) {
      annotateRatingChangeForAll(allGamesData); // Annotate before writing
      writeGamesToSheet(sheet, allGamesData);
      writeMovesMatrixSheets(allGamesData);
      writeDailySummaryFromArray(allGamesData);
      writeOpponentSummaryFromArray(allGamesData);
    }
    
    console.log(`Successfully fetched ${totalGames} games with complete data!`);
    
    // Show completion message
    SpreadsheetApp.getUi().alert(
      'Chess.com Complete Data Import',
      `Successfully imported ${totalGames} games with all available data fields!`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Error fetching Chess.com games:', error);
    SpreadsheetApp.getUi().alert(
      'Error',
      `Failed to fetch games: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Get list of game archives from Chess.com API
 */
function getGameArchives() {
  const url = `https://api.chess.com/pub/player/${CHESS_COM_USERNAME}/games/archives`;
  
  try {
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());
    
    if (!data.archives) {
      throw new Error('No game archives found. Check if username is correct.');
    }
    
    return data.archives;
  } catch (error) {
    throw new Error(`Failed to fetch game archives: ${error.message}`);
  }
}

/**
 * Fetch games from a specific monthly archive
 */
function fetchGamesFromArchive(archiveUrl) {
  try {
    const response = UrlFetchApp.fetch(archiveUrl);
    const data = JSON.parse(response.getContentText());
    
    if (!data.games) {
      return [];
    }
    
    // Process each game and extract ALL available data
    return data.games.map(game => processCompleteGameData(game));
  } catch (error) {
    console.error(`Error fetching archive ${archiveUrl}:`, error);
    return [];
  }
}

/**
 * Process individual game data extracting ALL possible fields
 */
function processCompleteGameData(game) {
  const isWhite = game.white.username.toLowerCase() === CHESS_COM_USERNAME.toLowerCase();
  const opponent = isWhite ? game.black : game.white;
  const me = isWhite ? game.white : game.black;
  
  // Determine result from my perspective
  let myResult = 'Draw';
  if ((game.white.result === 'win' && isWhite) || (game.black.result === 'win' && !isWhite)) {
    myResult = 'Win';
  } else if (game.white.result === 'win' || game.black.result === 'win') {
    myResult = 'Loss';
  }
  
  // Extract opening information from PGN
  const openingInfo = extractOpeningInfo(game.pgn);
  
  // Extract detailed PGN components
  const pgnComponents = parsePgnComponents(game.pgn);
  const gameLengthMinutes = computeGameLengthMinutes({
    format: determineFormat(game),
    pgnStartTime: pgnComponents.startTime,
    pgnEndTime: pgnComponents.endTime
  });
  
  return {
    // Basic Game Information
    gameUrl: game.url || '',
    gameId: extractGameId(game.url) || '',
    startTime: game.start_time ? new Date(game.start_time * 1000) : '',
    endTime: game.end_time ? new Date(game.end_time * 1000) : '',
    timeControl: game.time_control || '',
    timeClass: game.time_class || '',
    format: determineFormat(game),
    rules: game.rules || 'chess',
    rated: game.rated !== undefined ? (game.rated ? 'Yes' : 'No') : 'N/A',
    
    // Game Result Information
    myResult: myResult,
    myResultCode: me.result || '',
    opponentResultCode: opponent.result || '',
    
    // Player Information (Me)
    myColor: isWhite ? 'White' : 'Black',
    myUsername: me.username || '',
    myRating: me.rating || '',
    myRatingAfter: me.rating || '', // Same as rating in archives
    myPlayerId: me['@id'] || '',
    myUuid: me.uuid || '', // Available in live games
    
    // Player Information (Opponent)
    opponentColor: isWhite ? 'Black' : 'White',
    opponentUsername: opponent.username || '',
    opponentRating: opponent.rating || '',
    opponentRatingAfter: opponent.rating || '', // Same as rating in archives
    opponentPlayerId: opponent['@id'] || '',
    opponentUuid: opponent.uuid || '', // Available in live games
    
    // Accuracy Data
    myAccuracy: game.accuracies && game.accuracies[isWhite ? 'white' : 'black'] ? 
                (game.accuracies[isWhite ? 'white' : 'black'] * 100).toFixed(1) + '%' : '',
    opponentAccuracy: game.accuracies && game.accuracies[isWhite ? 'black' : 'white'] ? 
                      (game.accuracies[isWhite ? 'black' : 'white'] * 100).toFixed(1) + '%' : '',
    
    // Opening Information
    ecoCode: openingInfo.eco || '',
    ecoUrl: game.eco || '',
    openingName: openingInfo.opening || '',
    variation: openingInfo.variation || '',
    
    // Game Analysis
    finalFen: game.fen || '',
    moveCount: countMoves(game.pgn),
    gameLengthMinutes: gameLengthMinutes,
    
    // Tournament/Match Information
    tournamentUrl: game.tournament || '',
    tournamentName: extractTournamentName(game.tournament) || '',
    matchUrl: game.match || '',
    matchName: extractMatchName(game.match) || '',
    
    // Current Game Data (for ongoing games)
    currentTurn: game.turn || '',
    isMyTurn: game.turn === (isWhite ? 'white' : 'black') ? 'Yes' : 'No',
    moveBy: game.move_by ? new Date(game.move_by * 1000) : '',
    drawOffer: game.draw_offer || '',
    hasDrawOffer: game.draw_offer ? 'Yes' : 'No',
    lastActivity: game.last_activity ? new Date(game.last_activity * 1000) : '',
    
    // PGN Component Data
    pgnEvent: pgnComponents.event,
    pgnSite: pgnComponents.site,
    pgnDate: pgnComponents.date,
    pgnRound: pgnComponents.round,
    pgnWhite: pgnComponents.white,
    pgnBlack: pgnComponents.black,
    pgnResult: pgnComponents.result,
    pgnWhiteElo: pgnComponents.whiteElo,
    pgnBlackElo: pgnComponents.blackElo,
    pgnTimeControl: pgnComponents.timeControl,
    pgnTermination: pgnComponents.termination,
    pgnStartTime: pgnComponents.startTime,
    pgnEndTime: pgnComponents.endTime,
    pgnLink: pgnComponents.link,
    pgnCurrentPosition: pgnComponents.currentPosition,
    pgnTimezone: pgnComponents.timezone,
    pgnEcoCode: pgnComponents.ecoCode,
    pgnOpening: pgnComponents.opening,
    pgnVariation: pgnComponents.variation,
    pgnUtcDate: pgnComponents.utcDate,
    pgnUtcTime: pgnComponents.utcTime,
    pgnWhiteTitle: pgnComponents.whiteTitle,
    pgnBlackTitle: pgnComponents.blackTitle,
    pgnWhiteTeam: pgnComponents.whiteTeam,
    pgnBlackTeam: pgnComponents.blackTeam,
    movesSummary: pgnComponents.moves,
    
    // PGN Data
    fullPgn: game.pgn || '',
    pgnHeaders: extractPgnHeaders(game.pgn),
    
    // Additional Analysis
    gameType: determineGameType(game),
    isLiveGame: game.time_class && ['blitz', 'bullet', 'rapid'].includes(game.time_class) ? 'Yes' : 'No',
    isDailyGame: game.time_class === 'daily' ? 'Yes' : 'No',
    isVariant: game.rules !== 'chess' ? 'Yes' : 'No',
    
    // Time Analysis
    baseTime: extractBaseTime(game.time_control),
    increment: extractIncrement(game.time_control),
    timeControlCategory: categorizeTimeControl(game.time_control),
    
    // Raw Data (for advanced users)
    rawGameData: JSON.stringify(game)
  };
}

/**
 * Extract game ID from URL
 */
function extractGameId(url) {
  if (!url) return null;
  const match = url.match(/\/game\/(?:live|daily)\/(\d+)/i) || url.match(/\/game\/(\d+)/);
  return match ? match[1] : null;
}

/**
 * Extract comprehensive opening information from PGN
 */
function extractOpeningInfo(pgn) {
  if (!pgn) return { eco: null, opening: null, variation: null };
  
  const ecoMatch = pgn.match(/\[ECO "([^"]+)"\]/);
  const openingMatch = pgn.match(/\[Opening "([^"]+)"\]/);
  const variationMatch = pgn.match(/\[Variation "([^"]+)"\]/);
  
  return {
    eco: ecoMatch ? ecoMatch[1] : null,
    opening: openingMatch ? openingMatch[1] : null,
    variation: variationMatch ? variationMatch[1] : null
  };
}

/**
 * Count number of moves in the game
 */
function countMoves(pgn) {
  if (!pgn) return 0;
  
  // Extract the moves section (after the headers)
  const movesSection = pgn.split('\n\n')[1] || '';
  // Count move numbers (format: "1." "2." etc.)
  const moveNumbers = (movesSection.match(/\d+\./g) || []).length;
  return moveNumbers;
}

/**
 * Calculate game duration in minutes
 */
function calculateGameDuration(game) {
  if (!game.start_time || !game.end_time) return '';
  
  const durationSeconds = game.end_time - game.start_time;
  const durationMinutes = Math.round(durationSeconds / 60);
  
  if (durationMinutes < 60) {
    return `${durationMinutes}m`;
  } else {
    const hours = Math.floor(durationMinutes / 60);
    const minutes = durationMinutes % 60;
    return `${hours}h ${minutes}m`;
  }
}

/**
 * Calculate average move time
 */
function calculateAverageMoveTime(game) {
  if (!game.start_time || !game.end_time || !game.pgn) return '';
  
  const duration = game.end_time - game.start_time;
  const moves = countMoves(game.pgn);
  
  if (moves === 0) return '';
  
  const avgSeconds = Math.round(duration / (moves * 2)); // Total plies = moves * 2
  
  if (avgSeconds < 60) {
    return `${avgSeconds}s`;
  } else {
    const minutes = Math.floor(avgSeconds / 60);
    const seconds = avgSeconds % 60;
    return `${minutes}m ${seconds}s`;
  }
}

/**
 * Extract tournament name from URL
 */
function extractTournamentName(tournamentUrl) {
  if (!tournamentUrl) return null;
  
  // Extract from URL path
  const match = tournamentUrl.match(/tournament\/([^\/]+)/);
  if (match) {
    return match[1].replace(/-/g, ' ').replace(/\b\w/g, l => l.toUpperCase());
  }
  return null;
}

/**
 * Extract match name from URL
 */
function extractMatchName(matchUrl) {
  if (!matchUrl) return null;
  
  // Extract from URL path
  const match = matchUrl.match(/match\/(\d+)/);
  return match ? `Match ${match[1]}` : null;
}

/**
 * Extract PGN headers as a formatted string (remaining headers not in separate columns)
 */
function extractPgnHeaders(pgn) {
  if (!pgn) return '';
  
  const headers = [];
  const headerRegex = /\[([^"]+?)\s+"([^"]*?)"\]/g;
  let match;
  
  // List of headers that have their own columns
  const separateHeaders = [
    'Event', 'Site', 'Date', 'Round', 'White', 'Black', 'Result', 
    'WhiteElo', 'BlackElo', 'TimeControl', 'Termination', 'StartTime', 
    'EndTime', 'Link', 'CurrentPosition', 'Timezone', 'ECO', 'Opening', 
    'Variation', 'UTCDate', 'UTCTime', 'WhiteTitle', 'BlackTitle', 
    'WhiteTeam', 'BlackTeam'
  ];
  
  while ((match = headerRegex.exec(pgn)) !== null) {
    const key = match[1].trim();
    const value = match[2].trim();
    
    if (!separateHeaders.includes(key)) {
      headers.push(`${key}: ${value}`);
    }
  }
  
  return headers.join('; ');
}

/**
 * Parse PGN into structured components used by the sheet
 */
function parsePgnComponents(pgn) {
  const result = {
    event: '',
    site: '',
    date: '',
    round: '',
    white: '',
    black: '',
    result: '',
    whiteElo: '',
    blackElo: '',
    timeControl: '',
    termination: '',
    startTime: '',
    endTime: '',
    link: '',
    currentPosition: '',
    timezone: '',
    ecoCode: '',
    opening: '',
    variation: '',
    utcDate: '',
    utcTime: '',
    whiteTitle: '',
    blackTitle: '',
    whiteTeam: '',
    blackTeam: '',
    moves: ''
  };

  if (!pgn || typeof pgn !== 'string') {
    return result;
  }

  // Collect headers
  const headerRegex = /\[([^\s]+)\s+"([^"]*)"\]/g;
  let match;
  const headers = {};
  while ((match = headerRegex.exec(pgn)) !== null) {
    const key = match[1].toLowerCase();
    headers[key] = match[2];
  }

  result.event = headers['event'] || '';
  result.site = headers['site'] || '';
  result.date = headers['date'] || '';
  result.round = headers['round'] || '';
  result.white = headers['white'] || '';
  result.black = headers['black'] || '';
  result.result = headers['result'] || '';
  result.whiteElo = headers['whiteelo'] || '';
  result.blackElo = headers['blackelo'] || '';
  result.timeControl = headers['timecontrol'] || '';
  result.termination = headers['termination'] || '';
  result.startTime = headers['starttime'] || headers['utctime'] || '';
  result.endTime = headers['endtime'] || '';
  result.link = headers['link'] || '';
  result.currentPosition = headers['currentposition'] || '';
  result.timezone = headers['timezone'] || '';
  result.ecoCode = headers['eco'] || '';
  result.opening = headers['opening'] || '';
  result.variation = headers['variation'] || '';
  result.utcDate = headers['utcdate'] || '';
  result.utcTime = headers['utctime'] || '';
  result.whiteTitle = headers['whitetitle'] || '';
  result.blackTitle = headers['blacktitle'] || '';
  result.whiteTeam = headers['whiteteam'] || '';
  result.blackTeam = headers['blackteam'] || '';

  // Extract moves section and build keyed map for moves with clocks
  const blankLineIndex = pgn.indexOf('\n\n');
  if (blankLineIndex !== -1) {
    let movesText = pgn.slice(blankLineIndex + 2).trim();

    // Remove result token at the end if present
    movesText = movesText.replace(/\s+(1-0|0-1|1\/2-1\/2|\*)\s*$/, '');

    // Remove RAVs/variations in parentheses
    movesText = movesText.replace(/\([^)]*\)/g, '');

    // Keep comments for clock extraction. Remove NAGs $n
    movesText = movesText.replace(/\$\d+/g, '');

    const movesMap = {};

    // Match full move pairs like: 12. e4 {[%clk 0:03:00]} e5 {[%clk 0:02:59.9]}
    const pairRegex = /(\d+)\.\s*(?!\.\.\.)([^\s{}]+)(?:\s*\{([^}]*)\})?(?:\s+(?:(?:\d+)?\.{3}\s*)?([^\s{}]+)(?:\s*\{([^}]*)\})?)?/g;
    let m;
    while ((m = pairRegex.exec(movesText)) !== null) {
      const moveNo = m[1];
      const whiteSan = m[2];
      const whiteComment = m[3] || '';
      const blackSan = m[4];
      const blackComment = m[5] || '';

      const wKey = `${moveNo}w`;
      if (whiteSan && whiteSan !== '...' && whiteSan !== '.') {
        const wClkMatch = whiteComment.match(/\[%clk\s+([0-9:\.]+)\]/);
        movesMap[wKey] = [whiteSan, wClkMatch ? wClkMatch[1] : ''];
      }

      if (blackSan && blackSan !== '...' && blackSan !== '.') {
        const bKey = `${moveNo}b`;
        const bClkMatch = (blackComment || '').match(/\[%clk\s+([0-9:\.]+)\]/);
        movesMap[bKey] = [blackSan, bClkMatch ? bClkMatch[1] : ''];
      }
    }

    // Handle black-only notation like: 23... c5 {[%clk 0:00:42]}
    const blackOnlyRegex = /(\d+)\.\s*\.\.\.\s*([^\s{}]+)(?:\s*\{([^}]*)\})?/g;
    while ((m = blackOnlyRegex.exec(movesText)) !== null) {
      const moveNo = m[1];
      const blackSan = m[2];
      const blackComment = m[3] || '';
      const bKey = `${moveNo}b`;
      if (!movesMap[bKey] && blackSan !== '.' && blackSan !== '...') {
        const bClkMatch = blackComment.match(/\[%clk\s+([0-9:\.]+)\]/);
        movesMap[bKey] = [blackSan, bClkMatch ? bClkMatch[1] : ''];
      }
    }

    result.moves = JSON.stringify(movesMap);
  }

  return result;
}

/**
 * Determine game type based on various factors
 */
function determineGameType(game) {
  if (game.tournament) return 'Tournament';
  if (game.match) return 'Team Match';
  if (game.rules !== 'chess') return `Variant (${game.rules})`;
  if (game.time_class === 'daily') return 'Daily';
  return 'Casual';
}

/**
 * Determine rating format bucket per user rules
 */
function determineFormat(game) {
  const rules = (game.rules || 'chess').toLowerCase();
  const timeClass = (game.time_class || '').toLowerCase();

  // Standard chess: mapped to time class or Daily
  if (rules === 'chess') {
    if (timeClass === 'daily') return 'Daily';
    if (timeClass === 'bullet') return 'Bullet';
    if (timeClass === 'blitz') return 'Blitz';
    if (timeClass === 'rapid') return 'Rapid';
    return timeClass ? timeClass.charAt(0).toUpperCase() + timeClass.slice(1) : '';
  }

  // Chess960 special cases
  if (rules === 'chess960') {
    return timeClass === 'daily' ? 'Daily 960' : 'Live 960';
  }

  // Other variants: format is just the variant name (title-cased)
  return rules
    .replace(/[-_]+/g, ' ')
    .replace(/\b\w/g, c => c.toUpperCase());
}

/**
 * Extract base time from time control
 */
function extractBaseTime(timeControl) {
  if (!timeControl) return '';
  
  // Daily format like "1/86400" -> not applicable
  if (timeControl.includes('/')) return '';
 
  const basePart = timeControl.split('+')[0];
  const seconds = parseInt(basePart, 10);
  if (isNaN(seconds)) return '';
  return seconds / 60; // minutes as a numeric value
}

/**
 * Extract increment from time control
 */
function extractIncrement(timeControl) {
  if (!timeControl) return '';
  
  // Daily format like "1/86400" -> not applicable
  if (timeControl.includes('/')) return '';
 
  const parts = timeControl.split('+');
  if (parts.length < 2) return 0;
  const inc = parseInt(parts[1], 10);
  return isNaN(inc) ? 0 : inc; // seconds as a numeric value
}

/**
 * Categorize time control
 */
function categorizeTimeControl(timeControl) {
  if (!timeControl) return '';
  
  if (timeControl.includes('/')) return 'Daily';
  
  const match = timeControl.match(/^(\d+)/);
  if (match) {
    const seconds = parseInt(match[1]);
    if (seconds < 180) return 'Bullet';
    if (seconds < 600) return 'Blitz';
    if (seconds < 1800) return 'Rapid';
    return 'Classical';
  }
  
  return 'Unknown';
}

/**
 * Get or create the target sheet
 */
function getOrCreateSheet() {
  let spreadsheet;
  
  try {
    spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  } catch (error) {
    throw new Error(`Cannot open spreadsheet with ID: ${SHEET_ID}. Please check the Sheet ID.`);
  }
  
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    console.log(`Created new sheet: ${SHEET_NAME}`);
  }
  
  return sheet;
}

/**
 * Set up comprehensive column headers
 */
function setupHeaders(sheet) {
  const headers = [
    // Basic Game Information
    'Game URL', 'Game ID', 'Start Time', 'End Time', 'Time Control', 'Time Class', 'Format', 
    'Rules', 'Rated', 'Game Type', 'Is Live Game', 'Is Daily Game', 'Is Variant',
    
    // Results
    'My Result', 'My Result Code', 'Opponent Result Code',
    
    // My Information
    'My Color', 'My Username', 'My Rating', 'Rating Change', 'My Player ID', 'My UUID',
    
    // Opponent Information  
    'Opponent Color', 'Opponent Username', 'Opponent Rating', 'Opponent Player ID', 'Opponent UUID',
    
    // Accuracy
    'My Accuracy', 'Opponent Accuracy',
    
    // Opening Information
    'ECO Code', 'ECO URL', 'Opening Name', 'Variation',
    
    // Game Analysis
    'Final FEN', 'Move Count', 'Game Length (min)',
    
    // Time Control Details
    'Base Time', 'Increment', 'Time Control Category',
    
    // Tournament/Match
    'Tournament URL', 'Tournament Name', 'Match URL', 'Match Name',
    
    // Current Game Status
    'Current Turn', 'Is My Turn', 'Move By', 'Draw Offer', 'Has Draw Offer', 'Last Activity',
    
    // PGN Component Headers
    'PGN Event', 'PGN Site', 'PGN Date', 'PGN Round', 'PGN White', 'PGN Black', 'PGN Result',
    'PGN White Elo', 'PGN Black Elo', 'PGN Time Control', 'PGN Termination', 'PGN Start Time',
    'PGN End Time', 'PGN Link', 'PGN Current Position', 'PGN Timezone', 'PGN ECO Code',
    'PGN Opening', 'PGN Variation', 'PGN UTC Date', 'PGN UTC Time', 'PGN White Title', 'PGN Black Title', 'PGN White Team', 'PGN Black Team', 'Moves Summary',
    
    // Complete Game Data
    'Full PGN', 'Raw Game Data'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Remove all formatting features (no-op)
}

/**
 * Write complete games data to the sheet
 */
function writeGamesToSheet(sheet, gamesData) {
  if (gamesData.length === 0) return;
  
  // Convert games data to 2D array for sheet
  const rows = gamesData.map(buildGameRowValues);
  
  // Write data starting from row 2 (after headers)
  const range = sheet.getRange(2, 1, rows.length, rows[0].length);
  range.setValues(rows);
  
  // No formatting
}

/**
 * Write matrices for moves: one sheet for SAN, one for clocks
 * Each game is a row; columns are 1w, 1b, 2w, 2b, ... up to the maximum move number across all games
 */
function writeMovesMatrixSheets(gamesData) {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  let sanSheet = spreadsheet.getSheetByName(MOVES_SAN_SHEET_NAME);
  if (!sanSheet) sanSheet = spreadsheet.insertSheet(MOVES_SAN_SHEET_NAME);
  else sanSheet.clear();
  let clkSheet = spreadsheet.getSheetByName(MOVES_CLOCK_SHEET_NAME);
  if (!clkSheet) clkSheet = spreadsheet.insertSheet(MOVES_CLOCK_SHEET_NAME);
  else clkSheet.clear();

  // Determine maximum move number across all games
  let maxMoveNo = 0;
  const parsedMovesByGame = [];
  for (let i = 0; i < gamesData.length; i++) {
    let movesObj = {};
    try {
      movesObj = gamesData[i].movesSummary ? JSON.parse(gamesData[i].movesSummary) : {};
    } catch (e) {
      movesObj = {};
    }
    parsedMovesByGame.push(movesObj);
    for (const key in movesObj) {
      if (!Object.prototype.hasOwnProperty.call(movesObj, key)) continue;
      const n = parseInt(key, 10);
      if (!isNaN(n) && n > maxMoveNo) maxMoveNo = n;
    }
  }

  // Build headers
  const headers = ['Game ID'];
  for (let n = 1; n <= maxMoveNo; n++) {
    headers.push(`${n}w`, `${n}b`);
  }
  sanSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  clkSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Build rows
  const sanRows = [];
  const clkRows = [];
  for (let i = 0; i < gamesData.length; i++) {
    const game = gamesData[i];
    const movesObj = parsedMovesByGame[i] || {};
    const sanRow = new Array(headers.length).fill('');
    const clkRow = new Array(headers.length).fill('');
    sanRow[0] = game.gameId || '';
    clkRow[0] = game.gameId || '';
    for (let n = 1; n <= maxMoveNo; n++) {
      const wKey = `${n}w`;
      const bKey = `${n}b`;
      const wVal = movesObj[wKey];
      const bVal = movesObj[bKey];
      // Column index: 1-based for headers; position in array is same as headers index
      const wColIdx = (n - 1) * 2 + 1; // after Game ID
      const bColIdx = (n - 1) * 2 + 2;
      if (Array.isArray(wVal)) {
        sanRow[wColIdx] = wVal[0] || '';
        clkRow[wColIdx] = wVal[1] || '';
      } else if (typeof wVal === 'string') {
        sanRow[wColIdx] = wVal;
      }
      if (Array.isArray(bVal)) {
        sanRow[bColIdx] = bVal[0] || '';
        clkRow[bColIdx] = bVal[1] || '';
      } else if (typeof bVal === 'string') {
        sanRow[bColIdx] = bVal;
      }
    }
    sanRows.push(sanRow);
    clkRows.push(clkRow);
  }

  if (sanRows.length > 0) sanSheet.getRange(2, 1, sanRows.length, headers.length).setValues(sanRows);
  if (clkRows.length > 0) clkSheet.getRange(2, 1, clkRows.length, headers.length).setValues(clkRows);
}

/**
 * Apply comprehensive formatting to the sheet
 */
function formatSheet(sheet, dataRows) {
  // Intentionally left blank: no formatting (row colors, sizes, borders, resizing, or number formats)
}

/**
 * Create a menu in the Google Sheet for easy access
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Chess.com Complete Import')
    .addItem('Fetch All Games (Complete Data)', 'fetchAllChessComGames')
    .addItem('Add New Games (Incremental)', 'addNewGames')
    .addToUi();
}

/**
 * Build a single row array for the Games sheet from a processed game record
 */
function buildGameRowValues(game) {
  return [
    // Basic Game Information
    game.gameUrl, game.gameId, game.startTime, game.endTime, game.timeControl, 
    game.timeClass, game.format, game.rules, game.rated, game.gameType, game.isLiveGame, 
    game.isDailyGame, game.isVariant,
    
    // Results
    game.myResult, game.myResultCode, game.opponentResultCode,
    
    // My Information
    game.myColor, game.myUsername, game.myRating, game.ratingChange !== undefined ? game.ratingChange : '', game.myPlayerId, game.myUuid,
    
    // Opponent Information
    game.opponentColor, game.opponentUsername, game.opponentRating, 
    game.opponentPlayerId, game.opponentUuid,
    
    // Accuracy
    game.myAccuracy, game.opponentAccuracy,
    
    // Opening Information
    game.ecoCode, game.ecoUrl, game.openingName, game.variation,
    
    // Game Analysis
    game.finalFen, game.moveCount, game.gameLengthMinutes,
    
    // Time Control Details
    game.baseTime, game.increment, game.timeControlCategory,
    
    // Tournament/Match
    game.tournamentUrl, game.tournamentName, game.matchUrl, game.matchName,
    
    // Current Game Status
    game.currentTurn, game.isMyTurn, game.moveBy, game.drawOffer, 
    game.hasDrawOffer, game.lastActivity,
    
    // PGN Component Headers
    game.pgnEvent, game.pgnSite, game.pgnDate, game.pgnRound, game.pgnWhite, 
    game.pgnBlack, game.pgnResult, game.pgnWhiteElo, game.pgnBlackElo, 
    game.pgnTimeControl, game.pgnTermination, game.pgnStartTime, game.pgnEndTime, 
    game.pgnLink, game.pgnCurrentPosition, game.pgnTimezone, game.pgnEcoCode, 
    game.pgnOpening, game.pgnVariation, game.pgnUtcDate, game.pgnUtcTime, game.pgnWhiteTitle, game.pgnBlackTitle, game.pgnWhiteTeam, game.pgnBlackTeam, game.movesSummary,
    
    // Complete Game Data
    game.fullPgn, game.rawGameData
  ];
}

/**
 * Compute rating change for a full dataset (rebuild): per format, based on endTime order
 */
function annotateRatingChangeForAll(gamesData) {
  const formatToIndices = {};
  for (let i = 0; i < gamesData.length; i++) {
    const f = gamesData[i].format || '';
    if (!formatToIndices[f]) formatToIndices[f] = [];
    formatToIndices[f].push(i);
  }
  for (const f in formatToIndices) {
    const idxs = formatToIndices[f].slice();
    // Sort by numeric Game ID ascending; fallback to endTime
    idxs.sort((a, b) => {
      const ai = Number(gamesData[a].gameId);
      const bi = Number(gamesData[b].gameId);
      if (Number.isFinite(ai) && Number.isFinite(bi) && ai !== bi) return ai - bi;
      const ta = gamesData[a].endTime instanceof Date ? gamesData[a].endTime.getTime() : 0;
      const tb = gamesData[b].endTime instanceof Date ? gamesData[b].endTime.getTime() : 0;
      return ta - tb;
    });
    let last = null;
    for (let k = 0; k < idxs.length; k++) {
      const i = idxs[k];
      const r = Number(gamesData[i].myRating);
      if (Number.isFinite(r) && Number.isFinite(last)) {
        gamesData[i].ratingChange = r - last;
      } else {
        gamesData[i].ratingChange = '';
      }
      if (Number.isFinite(r)) last = r;
    }
  }
}

/**
 * For incremental add: compute rating change of new records using existing sheet history by format
 */
function annotateRatingChangeForNew(sheet, newGameRecords) {
  if (newGameRecords.length === 0) return;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol === 0) {
    newGameRecords.forEach(rec => rec.ratingChange = '');
    return;
  }
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const formatCol = headers.indexOf('Format') + 1;
  const idCol = headers.indexOf('Game ID') + 1;
  const myRatingCol = headers.indexOf('My Rating') + 1;
  if (!formatCol || !idCol || !myRatingCol) {
    newGameRecords.forEach(rec => rec.ratingChange = '');
    return;
  }
  const existing = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const byFormat = {};
  for (let i = 0; i < existing.length; i++) {
    const f = existing[i][formatCol - 1];
    const idStr = existing[i][idCol - 1];
    const id = Number(idStr);
    const r = Number(existing[i][myRatingCol - 1]);
    if (!byFormat[f]) byFormat[f] = [];
    byFormat[f].push({ id: Number.isFinite(id)? id:null, r: Number.isFinite(r)? r:null });
  }
  for (const f in byFormat) {
    byFormat[f].sort((a, b) => (a.id||0) - (b.id||0));
  }
  for (let i = 0; i < newGameRecords.length; i++) {
    const rec = newGameRecords[i];
    const f = rec.format || '';
    const id = Number(rec.gameId);
    const r = Number(rec.myRating);
    const list = byFormat[f] || [];
    // find largest id < current id
    let prev = null;
    for (let j = list.length - 1; j >= 0; j--) {
      if (Number.isFinite(list[j].id) && list[j].id < id) { prev = list[j]; break; }
    }
    if (prev && Number.isFinite(r) && Number.isFinite(prev.r)) {
      rec.ratingChange = r - prev.r;
    } else {
      rec.ratingChange = '';
    }
    // push current for subsequent new records
    list.push({ id: Number.isFinite(id)? id:null, r: Number.isFinite(r)? r:null });
    list.sort((a,b)=> (a.id||0)-(b.id||0));
    byFormat[f] = list;
  }
}

/**
 * Incrementally add new games only; insert at top of sheets
 */
function addNewGames() {
  const sheet = getOrCreateSheet();
  const existingIds = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    // Find Game ID column by header name
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idColIndex = headers.indexOf('Game ID') + 1; // 1-based
    if (idColIndex > 0) {
      const idValues = sheet.getRange(2, idColIndex, lastRow - 1, 1).getValues();
      for (let i = 0; i < idValues.length; i++) {
        const val = idValues[i][0];
        if (val) existingIds.add(String(val));
      }
    }
  }

  const archives = getGameArchives();
  const reversed = archives.slice().reverse(); // newest first
  const newGameRecords = [];

  for (let i = 0; i < reversed.length; i++) {
    const url = reversed[i];
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());
    const games = (data && data.games) ? data.games.slice() : [];
    // newest first by end_time
    games.sort((a, b) => (b.end_time || 0) - (a.end_time || 0));
    let newInThisArchive = 0;
    for (let g = 0; g < games.length; g++) {
      const game = games[g];
      const id = extractGameId(game.url);
      if (!id || existingIds.has(String(id))) {
        continue;
      }
      const record = processCompleteGameData(game);
      newGameRecords.push(record);
      existingIds.add(String(id));
      newInThisArchive++;
    }
    // If no new games found in this recent archive, assume older ones are already imported
    if (newInThisArchive === 0) {
      break;
    }
    // Consider a brief delay to reduce rate limiting
    Utilities.sleep(200);
  }

  if (newGameRecords.length === 0) {
    console.log('No new games to add.');
    return;
  }
  
  // Compute rating change for the new records using existing history, before writing
  annotateRatingChangeForNew(sheet, newGameRecords);
  // Insert into Games sheet at top (row 2)
  const rows = newGameRecords.map(buildGameRowValues);
  sheet.insertRows(2, rows.length);
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  
  // Update moves matrices incrementally
  prependMovesMatrixRows(newGameRecords);
  // Rebuild daily summary and opponents from the updated sheet
  writeDailySummary();
  writeOpponentSummary();
}

/**
 * Prepend new rows to Moves SAN and Moves Clock sheets; extend columns if needed
 */
function prependMovesMatrixRows(newGameRecords) {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  let sanSheet = spreadsheet.getSheetByName(MOVES_SAN_SHEET_NAME);
  if (!sanSheet) sanSheet = spreadsheet.insertSheet(MOVES_SAN_SHEET_NAME);
  let clkSheet = spreadsheet.getSheetByName(MOVES_CLOCK_SHEET_NAME);
  if (!clkSheet) clkSheet = spreadsheet.insertSheet(MOVES_CLOCK_SHEET_NAME);

  // Ensure headers exist
  if (sanSheet.getLastRow() === 0) {
    sanSheet.getRange(1, 1, 1, 1).setValues([['Game ID']]);
  }
  if (clkSheet.getLastRow() === 0) {
    clkSheet.getRange(1, 1, 1, 1).setValues([['Game ID']]);
  }

  const sanHeaders = sanSheet.getRange(1, 1, 1, sanSheet.getLastColumn() || 1).getValues()[0];
  const clkHeaders = clkSheet.getRange(1, 1, 1, clkSheet.getLastColumn() || 1).getValues()[0];
  const currentHeaderLen = Math.max(sanHeaders.length, clkHeaders.length);

  // Determine maximum move number needed from new records
  let neededMaxMoveNo = 0;
  const parsedMoves = [];
  for (let i = 0; i < newGameRecords.length; i++) {
    let obj = {};
    try {
      obj = newGameRecords[i].movesSummary ? JSON.parse(newGameRecords[i].movesSummary) : {};
    } catch (e) { obj = {}; }
    parsedMoves.push(obj);
    for (const k in obj) {
      const n = parseInt(k, 10);
      if (!isNaN(n) && n > neededMaxMoveNo) neededMaxMoveNo = n;
    }
  }

  // Compute desired header length: 1 (Game ID) + neededMaxMoveNo * 2
  const desiredHeaderLen = 1 + neededMaxMoveNo * 2;
  if (desiredHeaderLen > currentHeaderLen) {
    const newHeaders = ['Game ID'];
    for (let n = 1; n <= neededMaxMoveNo; n++) {
      newHeaders.push(`${n}w`, `${n}b`);
    }
    sanSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    clkSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  }

  const finalHeaderLen = Math.max(
    sanSheet.getLastColumn() || 1,
    clkSheet.getLastColumn() || 1
  );

  // Build new rows
  const sanRows = [];
  const clkRows = [];
  for (let i = 0; i < newGameRecords.length; i++) {
    const rec = newGameRecords[i];
    const movesObj = parsedMoves[i] || {};
    const sanRow = new Array(finalHeaderLen).fill('');
    const clkRow = new Array(finalHeaderLen).fill('');
    sanRow[0] = rec.gameId || '';
    clkRow[0] = rec.gameId || '';
    for (const key in movesObj) {
      const val = movesObj[key];
      const n = parseInt(key, 10);
      if (isNaN(n)) continue;
      const isWhite = key.endsWith('w');
      const colIdx = 1 + (n - 1) * 2 + (isWhite ? 0 : 1); // 0-based index
      if (colIdx >= finalHeaderLen) continue;
      if (Array.isArray(val)) {
        sanRow[colIdx] = val[0] || '';
        clkRow[colIdx] = val[1] || '';
      } else if (typeof val === 'string') {
        sanRow[colIdx] = val;
      }
    }
    sanRows.push(sanRow);
    clkRows.push(clkRow);
  }

  // Insert rows at top (after header)
  sanSheet.insertRows(2, sanRows.length);
  clkSheet.insertRows(2, clkRows.length);
  sanSheet.getRange(2, 1, sanRows.length, sanRows[0].length).setValues(sanRows);
  clkSheet.getRange(2, 1, clkRows.length, clkRows[0].length).setValues(clkRows);
}

/**
 * Compute game length in minutes from PGN StartTime/EndTime for live games only
 */
function computeGameLengthMinutes(record) {
  const liveFormats = ['Bullet','Blitz','Rapid'];
  if (liveFormats.indexOf(record.format) === -1) return 0;
  const startStr = record.pgnStartTime;
  const endStr = record.pgnEndTime;
  if (!startStr || !endStr) return 0;
  const startSec = parseDailyClockToSeconds(startStr);
  const endSec = parseDailyClockToSeconds(endStr);
  if (startSec == null || endSec == null) return 0;
  let delta = endSec - startSec;
  if (delta < 0) delta += 24 * 3600; // wrap midnight
  return Math.round((delta / 60) * 10) / 10;
}

function parseDailyClockToSeconds(hms) {
  const parts = String(hms).trim().split(':');
  if (parts.length < 2 || parts.length > 3) return null;
  const h = parseInt(parts[0], 10) || 0;
  const m = parseInt(parts[1], 10) || 0;
  const s = parts.length === 3 ? parseFloat(parts[2]) || 0 : 0;
  return h * 3600 + m * 60 + s;
}

/**
 * Build daily summary from an in-memory array of game records (for rebuild)
 */
function writeDailySummaryFromArray(gamesData) {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const sheetName = 'Daily';
  let dailySheet = spreadsheet.getSheetByName(sheetName);
  if (!dailySheet) dailySheet = spreadsheet.insertSheet(sheetName);
  dailySheet.clear();
  const rows = buildDailySummaryRows(gamesData);
  if (rows.length === 0) return;
  dailySheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}

/**
 * Build daily summary by reading from Games sheet (for incremental)
 */
function writeDailySummary() {
  const sheet = getOrCreateSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const idx = {};
  ['End Time','Format','My Rating','My Result','PGN Start Time','PGN End Time'].forEach(h => idx[h] = headers.indexOf(h));
  const games = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const format = r[idx['Format']];
    const endTime = r[idx['End Time']];
    const rating = Number(r[idx['My Rating']]);
    const result = r[idx['My Result']];
    const pStart = r[idx['PGN Start Time']];
    const pEnd = r[idx['PGN End Time']];
    const rec = {
      format: format,
      endTime: endTime,
      myRating: rating,
      myResult: result,
      pgnStartTime: pStart,
      pgnEndTime: pEnd
    };
    games.push(rec);
  }
  const dataRows = buildDailySummaryRows(games);
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const sheetName = 'Daily';
  let dailySheet = spreadsheet.getSheetByName(sheetName);
  if (!dailySheet) dailySheet = spreadsheet.insertSheet(sheetName);
  dailySheet.clear();
  if (dataRows.length > 0) dailySheet.getRange(1, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
}

/**
 * Return 2D array for daily summary with requested columns
 */
function buildDailySummaryRows(gamesArray) {
  const formats = ['Bullet','Blitz','Rapid'];
  // Aggregate by date (yyyy-MM-dd) and compute per-game seconds
  const byDay = {};
  let minTime = null, maxTime = null;
  for (let i = 0; i < gamesArray.length; i++) {
    const g = gamesArray[i];
    const d = g.endTime instanceof Date ? g.endTime : null;
    if (!d) continue;
    const timeMs = d.getTime();
    if (minTime === null || timeMs < minTime) minTime = timeMs;
    if (maxTime === null || timeMs > maxTime) maxTime = timeMs;
    const dayKey = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (!byDay[dayKey]) byDay[dayKey] = [];
    const startStr = g.pgnStartTime;
    const endStr = g.pgnEndTime;
    let lenSec = 0;
    if (formats.indexOf(g.format) !== -1 && startStr && endStr) {
      const s = parseDailyClockToSeconds(startStr);
      const e = parseDailyClockToSeconds(endStr);
      if (s != null && e != null) {
        lenSec = e - s; if (lenSec < 0) lenSec += 86400;
      }
    }
    const clone = Object.assign({}, g);
    clone._lengthSec = lenSec;
    byDay[dayKey].push(clone);
  }
  // Build complete list of days newest -> oldest, carrying ratings forward
  const outDays = [];
  if (minTime !== null && maxTime !== null) {
    const oneDayMs = 24*3600*1000;
    // Start at max date, go down to min date
    let cur = new Date(maxTime);
    const minDate = new Date(minTime);
    cur.setHours(0,0,0,0);
    minDate.setHours(0,0,0,0);
    while (cur.getTime() >= minDate.getTime()) {
      const key = Utilities.formatDate(cur, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      outDays.push(key);
      cur = new Date(cur.getTime() - oneDayMs);
    }
  }
  const header = [
    'Date',
    'Bullet Rating', 'Blitz Rating', 'Rapid Rating',
    'Games', 'Wins', 'Losses', 'Draws', 'Win %', 'Score', 'Rating Sum', 'Rating Change', 'Total Minutes', 'Total Seconds',
    'Bullet Games','Bullet W','Bullet L','Bullet D','Bullet Win %','Bullet Score','Bullet Time (min)',
    'Blitz Games','Blitz W','Blitz L','Blitz D','Blitz Win %','Blitz Score','Blitz Time (min)',
    'Rapid Games','Rapid W','Rapid L','Rapid D','Rapid Win %','Rapid Score','Rapid Time (min)'
  ];
  const out = [header];
  // Carry-forward per-format last rating
  const carried = { Bullet:null, Blitz:null, Rapid:null };
  let prevRatingSum = null;
  for (let di = 0; di < outDays.length; di++) {
    const day = outDays[di];
    const items = (byDay[day] || []).slice().sort((a,b) => (a.endTime - b.endTime));
    const perFmt = {};
    formats.forEach(f => perFmt[f] = { lastRating: carried[f], games:0, w:0, l:0, d:0, timeMin:0, timeSec:0 });
    let totalGames = 0, totalW = 0, totalL = 0, totalD = 0, totalSec = 0;

    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      if (formats.indexOf(it.format) === -1) continue;
      const f = it.format;
      if (Number.isFinite(it.myRating)) perFmt[f].lastRating = it.myRating;
      perFmt[f].games++; totalGames++;
      if (it.myResult === 'Win') { perFmt[f].w++; totalW++; }
      else if (it.myResult === 'Loss') { perFmt[f].l++; totalL++; }
      else { perFmt[f].d++; totalD++; }
      totalSec += it._lengthSec || 0;
      perFmt[f].timeSec += it._lengthSec || 0;
    }
    // Update carried ratings
    formats.forEach(f => { carried[f] = perFmt[f].lastRating; });
    // Final ratings for the day are the per-format lastRatings (carried if no games)
    const bulletR = perFmt['Bullet'].lastRating;
    const blitzR = perFmt['Blitz'].lastRating;
    const rapidR = perFmt['Rapid'].lastRating;
    const ratingSum = [bulletR,blitzR,rapidR].reduce((s,v)=> s + (Number.isFinite(v)? v:0), 0);
    const ratingChange = (prevRatingSum==null) ? '' : (ratingSum - prevRatingSum);
    prevRatingSum = ratingSum;

    const winPct = totalGames ? ( (totalW + (totalD/2)) / totalGames ) : 0;
    const scoreStr = totalGames ? `${(totalW + totalD/2).toFixed(1)}/${totalGames}` : '';

    // Per-format time in minutes
    formats.forEach(f => { perFmt[f].timeMin = Math.round((perFmt[f].timeSec/60)*10)/10; });

    const totalMinutes = Math.floor(totalSec / 60);
    const totalSecondsRemainder = Math.round((totalSec - totalMinutes*60) * 10) / 10;

    const bullet = perFmt['Bullet'];
    const blitz = perFmt['Blitz'];
    const rapid = perFmt['Rapid'];

    const fmtPart = (o)=> [o.games, o.w, o.l, o.d, o.games? ((o.w + o.d/2)/o.games):0, o.games? `${(o.w + o.d/2).toFixed(1)}/${o.games}`:'', o.timeMin];

    out.push([
      day,
      bulletR || '', blitzR || '', rapidR || '',
      totalGames, totalW, totalL, totalD, winPct, scoreStr, ratingSum, ratingChange, totalMinutes, totalSecondsRemainder,
      ...fmtPart(bullet),
      ...fmtPart(blitz),
      ...fmtPart(rapid)
    ]);
  }
  return out;
}

function writeOpponentSummaryFromArray(gamesData) {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const sheetName = 'Opponent Summary';
  let oppSheet = spreadsheet.getSheetByName(sheetName);
  if (!oppSheet) oppSheet = spreadsheet.insertSheet(sheetName);
  oppSheet.clear();
  const rows = buildOpponentSummaryRows(gamesData);
  if (rows.length === 0) return;
  oppSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}

function writeOpponentSummary() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const gamesSheet = getOrCreateSheet();
  const lastRow = gamesSheet.getLastRow();
  if (lastRow < 2) return;
  const lastCol = gamesSheet.getLastColumn();
  const headers = gamesSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rows = gamesSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const idx = {};
  ['Opponent Username','Opponent Rating','My Result','Game ID','End Time','Format'].forEach(h => idx[h] = headers.indexOf(h));
  const games = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    games.push({
      opponentUsername: r[idx['Opponent Username']],
      opponentRating: r[idx['Opponent Rating']],
      myResult: r[idx['My Result']],
      gameId: r[idx['Game ID']],
      endTime: r[idx['End Time']],
      format: r[idx['Format']]
    });
  }
  const dataRows = buildOpponentSummaryRows(games);
  const sheetName = 'Opponent Summary';
  let oppSheet = spreadsheet.getSheetByName(sheetName);
  if (!oppSheet) oppSheet = spreadsheet.insertSheet(sheetName);
  oppSheet.clear();
  if (dataRows.length > 0) oppSheet.getRange(1, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
}

function buildOpponentSummaryRows(games) {
  // Aggregate per opponent username (case-insensitive key)
  const map = {};
  const keyFor = (u)=> (u || '').toLowerCase();
  for (let i = 0; i < games.length; i++) {
    const g = games[i];
    const key = keyFor(g.opponentUsername);
    if (!key) continue;
    if (!map[key]) map[key] = { username: g.opponentUsername, games:0, w:0, l:0, d:0, ids:[], lastSeenRating:null };
    map[key].games++;
    if (g.myResult === 'Win') map[key].w++; else if (g.myResult === 'Loss') map[key].l++; else map[key].d++;
    if (g.gameId) map[key].ids.push(String(g.gameId));
    const r = Number(g.opponentRating);
    if (Number.isFinite(r)) map[key].lastSeenRating = r; // last occurrence used
  }
  const opps = Object.values(map);
  // Fetch profiles (best-effort)
  for (let i = 0; i < opps.length; i++) {
    const u = opps[i].username;
    const prof = fetchChessComProfileSafe(u);
    if (prof) {
      opps[i].title = prof.title || '';
      opps[i].name = prof.name || '';
      opps[i].country = prof.country || '';
      opps[i].status = prof.status || '';
      opps[i].url = prof.url || '';
      opps[i].joined = prof.joined || '';
      opps[i].last_online = prof.last_online || '';
    }
    Utilities.sleep(80);
  }
  // Sort by games desc
  opps.sort((a,b) => b.games - a.games);
  const header = [
    'Opponent Username','Name','Title','Country','Status','Profile URL','Joined','Last Online',
    'Games','Wins','Losses','Draws','Win %','Score','Last Seen Rating','Game IDs'
  ];
  const out = [header];
  for (let i = 0; i < opps.length; i++) {
    const o = opps[i];
    const winPct = o.games ? ( (o.w + o.d/2) / o.games ) : 0;
    const scoreStr = o.games ? `${(o.w + o.d/2).toFixed(1)}/${o.games}` : '';
    out.push([
      o.username || '', o.name || '', o.title || '', o.country || '', o.status || '', o.url || '', 
      o.joined || '', o.last_online || '',
      o.games, o.w, o.l, o.d, winPct, scoreStr, o.lastSeenRating || '', o.ids.join(',')
    ]);
  }
  return out;
}

function fetchChessComProfileSafe(username) {
  try {
    if (!username) return null;
    const url = `https://api.chess.com/pub/player/${String(username).toLowerCase()}`;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;
    return JSON.parse(res.getContentText());
  } catch (e) {
    return null;
  }
}
