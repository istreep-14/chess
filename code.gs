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

  // Extract moves section (text after the blank line following headers)
  const blankLineIndex = pgn.indexOf('\n\n');
  if (blankLineIndex !== -1) {
    let movesText = pgn.slice(blankLineIndex + 2).trim();
    // Remove PGN comments {...} and NAGs $n
    movesText = movesText.replace(/\{[^}]*\}/g, '').replace(/\$\d+/g, '');
    // Normalize whitespace
    movesText = movesText.replace(/\s+/g, ' ').trim();
    if (movesText.length > 300) {
      movesText = movesText.slice(0, 297) + '...';
    }
    result.moves = movesText;
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
 * Extract base time from time control
 */
function extractBaseTime(timeControl) {
  if (!timeControl) return '';
  
  const match = timeControl.match(/^(\d+)/);
  if (match) {
    const seconds = parseInt(match[1]);
    if (seconds < 60) return `${seconds}s`;
    if (seconds < 3600) return `${Math.floor(seconds/60)}m`;
    return `${Math.floor(seconds/3600)}h`;
  }
  return timeControl;
}

/**
 * Extract increment from time control
 */
function extractIncrement(timeControl) {
  if (!timeControl) return '';
  
  const match = timeControl.match(/\+(\d+)$/);
  return match ? `+${match[1]}s` : '';
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
    'Game URL', 'Game ID', 'Start Time', 'End Time', 'Time Control', 'Time Class', 
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
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  headerRange.setWrap(true);
  
  // Freeze header row
  sheet.setFrozenRows(1);
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
    game.timeClass, game.rules, game.rated, game.gameType, game.isLiveGame, 
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
  
  // Format the data
  formatSheet(sheet, rows.length);
}

/**
 * Apply comprehensive formatting to the sheet
 */
function formatSheet(sheet, dataRows) {
  // Auto-resize columns
  sheet.autoResizeColumns(1, sheet.getLastColumn());
  
  // Format datetime columns
  const dateTimeColumns = [3, 4, 45, 48]; // Start Time, End Time, Move By, Last Activity
  dateTimeColumns.forEach(col => {
    if (col <= sheet.getLastColumn()) {
      const dateRange = sheet.getRange(2, col, dataRows, 1);
      dateRange.setNumberFormat('MM/dd/yyyy hh:mm:ss');
    }
  });
  
  // Add borders
  const dataRange = sheet.getRange(1, 1, dataRows + 1, sheet.getLastColumn());
  dataRange.setBorder(true, true, true, true, true, true);
  
  // Alternate row colors for better readability
  for (let i = 2; i <= dataRows + 1; i++) {
    if (i % 2 === 0) {
      sheet.getRange(i, 1, 1, sheet.getLastColumn()).setBackground('#f8f9fa');
    }
  }
  
  // Freeze first few columns for easier navigation
  sheet.setFrozenColumns(5);
  
  // Set specific column widths for better readability
  sheet.setColumnWidth(1, 200); // Game URL
  sheet.setColumnWidth(70, 400); // Full PGN
  sheet.setColumnWidth(71, 300); // Raw Game Data
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
