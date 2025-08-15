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
const SHEET_NAME = 'Sheet1'; // Name of the sheet tab
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
    
    // Write all games data to sheet
    if (allGamesData.length > 0) {
      writeGamesToSheet(sheet, allGamesData);
      writeMovesMatrixSheets(allGamesData);
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
    gameDuration: calculateGameDuration(game),
    averageMoveTime: calculateAverageMoveTime(game),
    
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
  const match = url.match(/\/game\/(\d+)/);
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
    const pairRegex = /(\d+)\.\s*(?!\.\.\.)([^\s{}]+)(?:\s*\{([^}]*)\})?(?:\s+([^\s{}]+)(?:\s*\{([^}]*)\})?)?/g;
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
    'My Color', 'My Username', 'My Rating', 'My Player ID', 'My UUID',
    
    // Opponent Information  
    'Opponent Color', 'Opponent Username', 'Opponent Rating', 'Opponent Player ID', 'Opponent UUID',
    
    // Accuracy
    'My Accuracy', 'Opponent Accuracy',
    
    // Opening Information
    'ECO Code', 'ECO URL', 'Opening Name', 'Variation',
    
    // Game Analysis
    'Final FEN', 'Move Count', 'Game Duration', 'Average Move Time',
    
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
    'PGN Opening', 'PGN Variation', 'Moves Summary',
    
    // Complete Game Data
    'PGN Other Headers', 'Full PGN', 'Raw Game Data'
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
  const rows = gamesData.map(game => [
    // Basic Game Information
    game.gameUrl, game.gameId, game.startTime, game.endTime, game.timeControl, 
    game.timeClass, game.format, game.rules, game.rated, game.gameType, game.isLiveGame, 
    game.isDailyGame, game.isVariant,
    
    // Results
    game.myResult, game.myResultCode, game.opponentResultCode,
    
    // My Information
    game.myColor, game.myUsername, game.myRating, game.myPlayerId, game.myUuid,
    
    // Opponent Information
    game.opponentColor, game.opponentUsername, game.opponentRating, 
    game.opponentPlayerId, game.opponentUuid,
    
    // Accuracy
    game.myAccuracy, game.opponentAccuracy,
    
    // Opening Information
    game.ecoCode, game.ecoUrl, game.openingName, game.variation,
    
    // Game Analysis
    game.finalFen, game.moveCount, game.gameDuration, game.averageMoveTime,
    
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
    game.pgnOpening, game.pgnVariation, game.movesSummary,
    
    // Complete Game Data
    game.pgnHeaders, game.fullPgn, game.rawGameData
  ]);
  
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
    .addToUi();
}
