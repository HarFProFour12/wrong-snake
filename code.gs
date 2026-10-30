const size = 16;
const board = 'game';
const wait_time = 300;
const max_time = 280000;
const startCol = 3;
const startRow = 6;

const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName(board);

let goldenAppleChance = 0.25;
const goldenApplePoints = 5;

function setup() {
  titles();
  drawBoard();
  setupModeCell();
  updateHighscores();
}

function titles() {
  sheet.getRange('A1:T1').merge()
  .setValue('SPREADSHEET SNAKE GAME')
  .setHorizontalAlignment('center')
  .setVerticalAlignment('middle')
  .setFontSize(24)
  .setFontWeight('bold')
  .setBackground('#111827')
  .setFontColor(getThemeColor());

  sheet.getRange('A2:T2').merge()
  .setValue('Made by Harry Fanouriakis')
  .setHorizontalAlignment('center')
  .setVerticalAlignment('middle')
  .setFontSize(14)
  .setFontWeight('normal')
  .setBackground(lighten(getThemeColor()))
  .setFontColor('#111827');

  sheet.getRange('A3:T3').merge()
  .setBackground('#111827');

  sheet.getRange('A5:T5').merge()
  .setBackground('#111827');

  sheet.getRange('A4:G4').merge()
  .setValue('SCORE: 0 | BEST: ' + getBestScore())
  .setHorizontalAlignment('center')
  .setVerticalAlignment('middle')
  .setFontSize(12)
  .setFontWeight('bold')
  .setBackground('#111827')
  .setFontColor(getThemeColor());
    
  sheet.getRange('H4:M4').merge()
  .setValue('START')
  .setHorizontalAlignment('center')
  .setVerticalAlignment('middle')
  .setFontSize(16)
  .setFontWeight('bold')
  .setBackground(getThemeColor())
  .setFontColor('#111827');

  sheet.getRange('N4:T4').merge()
  .setBackground('#111827');

  sheet.getRange('V2:W2').merge()
  .setValue('HIGHSCORES')
  .setHorizontalAlignment('center')
  .setVerticalAlignment('middle')
  .setFontSize(20)
  .setFontWeight('bold')
  .setBackground('#111827')
  .setFontColor(getThemeColor());

  sheet.getRange('V3:W3')
  .setValues([['PLAYER', 'SCORE']])
  .setHorizontalAlignment('center')
  .setFontSize(12)
  .setFontWeight('bold')
  .setFontColor(lighten(getThemeColor()))
  .setBackground('#111827');

  sheet.getRange('V4:W8')
  .setHorizontalAlignment('center')
  .setVerticalAlignment('middle')
  .setFontSize(10)
  .setFontWeight('bold')
  .setBackground(lighten(getThemeColor()))
  .setFontColor('#111827')
  .setBorder(
  true, true, true, true, true, true,
  '#111827',
  SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange('X1:Z1').merge()
  .setValue("HOW IT WORKS")
  .setHorizontalAlignment('center')
  .setVerticalAlignment('middle')
  .setFontSize(24)
  .setFontWeight('bold')
  .setBackground('#ffffff')
  .setFontColor(getThemeColor());

  sheet.getRange('X2:Z12').breakApart();

  sheet.getRange('X2:Z9').merge()
  .setValue(
    '   • Change direction with the arrow buttons\n' +
    '   • 🍎 = 1 point\n' +
    '   • ⭐ = 5 points and +5 length\n' +
    '   • 🐢 = slow motion, ✂️ = shorter tail\n' +
    '   • Hit the border, a wall or yourself = game over\n' +
    '   • Harder difficulty = faster, fewer ⭐\n' +
    '   • The snake speeds up as you score\n' +
    '   • Impossible: eaten apples become walls\n' +
    '   • Wrap-around mode: borders teleport you\n' +
    '   • A game lasts about 4.5 minutes max\n' +
    '   • Top 5 scores go in HIGHSCORES (level-specific)'
  )
  .setHorizontalAlignment('left')
  .setVerticalAlignment('top')
  .setWrap(true)
  .setFontSize(11)
  .setFontWeight('normal')
  .setBackground('#ffffff')
  .setFontColor('#111827');
}

function onOpen() {
  setup();
}

function boardColor(r, c) {
  // r and c are 1-based board coordinates
  return (r + c) % 2 === 0 ? '#111827' : '#151E2E';
}

function drawBoard() {
  sheet.setHiddenGridlines(true);

  sheet.setColumnWidths(1, 20, 25);
  sheet.setRowHeights(startRow, 20, 25);

  sheet.getRange(startRow, 1, 20, 20)
    .setBackground('#111827')
    .setBorder(false, false, false, false, false, false);

  const whiteRows = 19 - size;
  if (whiteRows > 0) {
    sheet.getRange(startRow + size + 1, 1, whiteRows, 20)
      .setBackground('#ffffff');
  }

  // the board itself
  sheet.getRange(startRow, startCol, size, size)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(12);

  paintEmptyBoard();
  drawBoardBorder();
}

function drawBoardBorder() {
  const range = sheet.getRange(startRow, startCol, size, size);

  range.setBorder(
    true, true, true, true, true, true,
    '#374151',
    SpreadsheetApp.BorderStyle.SOLID
  );

  range.setBorder(
    true, true, true, true, false, false,
    '#ffffff',
    SpreadsheetApp.BorderStyle.SOLID_THICK
  );
}

function getThemeColor() {
  return PropertiesService.getDocumentProperties()
    .getProperty('themeColor') || '#22C55E';
}

function paintEmptyBoard() {
  const bg = Array.from({ length: size }, (_, i) =>
    Array.from({ length: size }, (_, j) => boardColor(i + 1, j + 1))
  );

  sheet.getRange(startRow, startCol, size, size).setBackgrounds(bg);
}

function drawSnake(gameState) {
  if (!gameState || !gameState.snake) return;
  
  const theme = getThemeColor();
  const headColor = lighten(theme);

  const bg = Array.from({ length: size }, (_, i) =>
    Array.from({ length: size }, (_, j) => boardColor(i + 1, j + 1))
  );
  const vals = Array.from({ length: size }, () => Array(size).fill(''));

  (gameState.walls || []).forEach(w => {
    bg[w.r - 1][w.c - 1] = '#6B7280';
    vals[w.r - 1][w.c - 1] = '▪';
  });

  if (gameState.food) {
    vals[gameState.food.r - 1][gameState.food.c - 1] =
      gameState.food.golden ? '⭐' : '🍎';
  }

  if (gameState.power) {
    vals[gameState.power.r - 1][gameState.power.c - 1] =
      gameState.power.type === 'slow' ? '🐢' : '✂️';
  }
  gameState.snake.forEach((s, i) => {
    bg[s.r - 1][s.c - 1] = i === 0 ? headColor : theme;
  });

  sheet.getRange(startRow, startCol, size, size)
    .setBackgrounds(bg)
    .setValues(vals);

  SpreadsheetApp.flush();
}

function saveGameState(gameState) {
  PropertiesService.getDocumentProperties()
    .setProperty('snakeState', JSON.stringify(gameState));
}

function loadGameState() {
  const saved = PropertiesService.getDocumentProperties()
    .getProperty('snakeState');

  return saved ? JSON.parse(saved) : null;
}

function startGame() {
  setup();

  const gameState = {
    snake: [
      { r: 8, c: 6 },
      { r: 8, c: 5 },
      { r: 8, c: 4 }
    ],
    direction: "RIGHT",
    running: true,
    score: 0,
    best: getBestScore(),
    grow: 0,
    reason: '',
    food: null,
    walls: [],
    power: null,
    slowTicks: 0,
    wrap: isWrapMode(),
    difficulty: currentDifficulty(),
    level: 0,
  };

  const props = PropertiesService.getDocumentProperties();
  props.setProperty('snakeDir', 'RIGHT');
  props.setProperty('snakePaused', 'false');

  gameState.food = spawnApple(gameState.snake, gameState.walls);

  saveGameState(gameState);
  drawSnake(gameState);

  playGame();
}

function resetScores() {
  const scoresSheet = ss.getSheetByName('scores');

  if (!scoresSheet || scoresSheet.getLastRow() < 2) return;

  scoresSheet.getRange(
    2, 1,
    scoresSheet.getLastRow() - 1,
    4
  ).clearContent();

  updateHighscores();
}

function moveSnake() {
  const gameState = loadGameState();
  if (!gameState || !gameState.running) return;

  const dir = PropertiesService.getDocumentProperties()
    .getProperty('snakeDir') || gameState.direction;

  gameState.direction = dir;
  gameState.lastMoved = dir;

  const head = gameState.snake[0];
  const next = { r: head.r, c: head.c };

  if (dir === 'UP') next.r--;
  if (dir === 'DOWN') next.r++;
  if (dir === 'LEFT') next.c--;
  if (dir === 'RIGHT') next.c++;

  // border: wrap around or die
  if (gameState.wrap) {
    if (next.r < 1) next.r = size;
    if (next.r > size) next.r = 1;
    if (next.c < 1) next.c = size;
    if (next.c > size) next.c = 1;
  } else if (next.r < 1 || next.r > size || next.c < 1 || next.c > size) {
    gameState.running = false;
    gameState.reason = 'border';
    saveGameState(gameState);
    return;
  }

  // obstacle collision
  if (gameState.walls && gameState.walls.some(w => w.r === next.r && w.c === next.c)) {
    gameState.running = false;
    gameState.reason = 'wall';
    saveGameState(gameState);
    return;
  }

  const ateApple =
    gameState.food &&
    next.r === gameState.food.r &&
    next.c === gameState.food.c;

  // tail stays in place while eating or still growing
  const growing = ateApple || gameState.grow > 0;
  const body = growing ? gameState.snake : gameState.snake.slice(0, -1);

  if (body.some(s => s.r === next.r && s.c === next.c)) {
    gameState.running = false;
    gameState.reason = 'self';
    saveGameState(gameState);
    return;
  }

  gameState.snake.unshift(next);

  if (ateApple) {
    if (gameState.food.golden) {
      gameState.score += goldenApplePoints;
      gameState.grow += 4;   // this move adds 1, so 4 more makes 5 total
    } else {
      gameState.score += 1;
    }

    if (getDifficulty() === "Impossible") {
      gameState.walls.push({ r: gameState.food.r, c: gameState.food.c });
    }

    gameState.food = spawnApple(gameState.snake, gameState.walls);

    if (gameState.power && gameState.food &&
        gameState.power.r === gameState.food.r &&
        gameState.power.c === gameState.food.c) {
      gameState.power = null;
    }

    if (!gameState.food) {
      gameState.running = false;
      gameState.reason = 'full';
    }
  } else if (gameState.grow > 0) {
    gameState.grow--;        // keep the tail, so the snake gets longer
  } else {
    gameState.snake.pop();
  }

    // power-ups
  if (gameState.slowTicks > 0) gameState.slowTicks--;

  if (gameState.power) {
    const p = gameState.power;

    if (next.r === p.r && next.c === p.c) {
      gameState.score += 2;

      if (p.type === 'slow') {
        gameState.slowTicks = 25;
      } else {
        const cut = Math.min(2, gameState.snake.length - 3);
        if (cut > 0) gameState.snake.splice(-cut, cut);
      }
      gameState.power = null;
    } else {
      p.ttl--;
      if (p.ttl <= 0) gameState.power = null;
    }
  } else if (Math.random() < 0.04) {
    gameState.power = spawnPower(gameState);
  }

  // level obstacles: new walls every 10 points
  const newLevel = Math.floor(gameState.score / 10);
  if (newLevel > (gameState.level || 0)) {
    gameState.level = newLevel;
    addLevelWalls(gameState);
  }

  saveGameState(gameState);
  drawSnake(gameState);

  sheet.getRange("A4").setValue(
    'SCORE: ' + gameState.score +
    ' | BEST: ' + Math.max(gameState.best || 0, gameState.score)
  );
}

function playGame() {
  const startTime = Date.now();
  const props = PropertiesService.getDocumentProperties();

  while (Date.now() - startTime < max_time) {
    const gameState = loadGameState();
    if (!gameState || !gameState.running) break;

    if (props.getProperty('snakePaused') === 'true') {
      Utilities.sleep(200);
      continue;
    }

    moveSnake();

    const after = loadGameState();
    const slow = after && after.slowTicks > 0 ? 1.8 : 1;
    Utilities.sleep(getWaitTime(after ? after.score : 0) * slow);
  }

  const gameState = loadGameState();
  if (!gameState) return;

  if (gameState.running) {
    gameState.running = false;
    gameState.reason = 'time';
    saveGameState(gameState);
  }

  props.setProperty('snakePaused', 'false');
  sheet.getRange('H4').setValue('START');

  clearBoard();

  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Game Over!',
    reasonText(gameState.reason) +
    '\nYour score: ' + gameState.score +
    '\nEnter your name:',
    ui.ButtonSet.OK_CANCEL
  );

  let playerName = 'Player';
  if (response.getSelectedButton() === ui.Button.OK) {
    const enteredName = response.getResponseText().trim();
    if (enteredName !== '') playerName = enteredName;
  }

  saveScore(gameState.score, playerName, gameState.difficulty);
  sheet.getRange('A4:G4').setValue('SCORE: 0 | BEST: ' + getBestScore());
}

function clearBoard() {
  sheet.getRange(startRow, startCol, size, size).clearContent();

  paintEmptyBoard();
  drawBoardBorder();
}

function spawnApple(snake, walls = []) {
  let apple;
  let attempts = 0;

  do {
    apple = {
      r: Math.floor(Math.random() * size) + 1,
      c: Math.floor(Math.random() * size) + 1
    };
    attempts++;

    if (attempts > size * size * 2) {
      return null;
    }
  } while (
    snake.some(segment =>
      segment.r === apple.r && segment.c === apple.c
    ) ||
    walls.some(wall =>
      wall.r === apple.r && wall.c === apple.c
    )
  );

  apple.golden = Math.random() < goldenAppleChance;
  return apple;
}

function stopGame() {
  const gameState = loadGameState();

  if (!gameState || !gameState.running) return;

  gameState.running = false;
  gameState.reason = 'stopped';
  saveGameState(gameState);
}

function changeDirection(newDirection) {
  const gameState = loadGameState();
  if (!gameState || !gameState.running) return;

  const opposite = { 
    UP: 'DOWN', 
    DOWN: 'UP', 
    LEFT: 'RIGHT', 
    RIGHT: 'LEFT' 
  };

  const lastMoved = gameState.lastMoved || gameState.direction;
  if (opposite[lastMoved] === newDirection) return;

  PropertiesService.getDocumentProperties()
    .setProperty('snakeDir', newDirection);
}

function goUp() {
  changeDirection('UP');
}

function goDown() {
  changeDirection('DOWN');
}

function goLeft() {
  changeDirection('LEFT');
}

function goRight() {
  changeDirection('RIGHT');
}

function themeGreen() {
  changeTheme('#22C55E');
}

function themeBlue() {
  changeTheme('#3B82F6');
}

function themePurple() {
  changeTheme('#A855F7');
}

function themePink() {
  changeTheme('#EC4899');
}

function themeCyan() {
  changeTheme('#06B6D4');
}

function themeOrange() {
  changeTheme('#FB923C');
}

function themeLime() {
  changeTheme('#A3E635');
}

function themeTeal() {
  changeTheme('#14B8A6');
}

function themeIndigo() {
  changeTheme('#6366F1');
}

function themeLavender() {
  changeTheme('#C4B5FD');
}

function changeTheme(color) {
  PropertiesService.getDocumentProperties()
    .setProperty('themeColor', color);

  titles();
  setupModeCell();

  const gameState = loadGameState();

  if (gameState && gameState.running) {
    drawSnake(gameState);
  }
}

function getDifficulty() {
  const difficulty = sheet.getRange("N4").getValue();
  console.log(difficulty);
  return difficulty;
}

function getWaitTime(score = 0) {
  const difficulty = sheet.getRange("N4").getValue();
  let base;

  if (difficulty === "Easy") {
    goldenAppleChance = 0.4;
    base = 600;
  } else if (difficulty === "Medium") {
    goldenAppleChance = 0.3;
    base = 450;
  } else if (difficulty === "Hard") {
    goldenAppleChance = 0.2;
    base = 300;
  } else if (difficulty === "Impossible") {
    goldenAppleChance = 0.1;
    base = 200;
  } else {
    goldenAppleChance = 0.25;
    base = 600;
  }

  return Math.max(100, base - Math.floor(score / 5) * 20);
}

function setupScores() {
  let scoresSheet = ss.getSheetByName('scores');

  if (!scoresSheet) {
    scoresSheet = ss.insertSheet('scores');
  }

  if (scoresSheet.getLastRow() === 0) {
    scoresSheet.appendRow(['Player', 'Score', 'Date', 'Difficulty']);
  }

  scoresSheet.getRange('D1').setValue('Difficulty');

  scoresSheet.getRange('A1:D1')
    .setFontWeight('bold')
    .setBackground('#22C55E')
    .setFontColor('#111827');
}

function saveScore(score, playerName, difficulty) {
  setupScores();

  const scoresSheet = ss.getSheetByName('scores');

  scoresSheet.appendRow([
    playerName,
    score,
    new Date(),
    difficulty || 'Medium'
  ]);

  scoresSheet.getRange(
    scoresSheet.getLastRow(), 3
  ).setNumberFormat('dd/MM/yyyy HH:mm:ss');

  updateHighscores();
}

function currentDifficulty() {
  return sheet.getRange('N4').getValue() || 'Medium';
}

function getScoresForDifficulty() {
  const s = ss.getSheetByName('scores');
  if (!s || s.getLastRow() < 2) return [];

  const diff = currentDifficulty();

  return s.getRange(2, 1, s.getLastRow() - 1, 4)
    .getValues()
    .filter(row => typeof row[1] === 'number' && row[3] === diff)
    .sort((a, b) => b[1] - a[1]);
}

function updateHighscores() {
  const top = getScoresForDifficulty()
    .slice(0, 5)
    .map(row => [row[0], row[1]]);

  sheet.getRange('V4:W8').clearContent();

  if (top.length > 0) {
    sheet.getRange(4, 22, top.length, 2).setValues(top);
  }
}

function getBestScore() {
  const scores = getScoresForDifficulty();
  return scores.length ? scores[0][1] : 0;
}

function lighten(hex) {
  const n = parseInt(hex.slice(1), 16);
  const mix = v => Math.round(v + (255 - v) * 0.55);
  return '#' + [mix(n >> 16), mix((n >> 8) & 255), mix(n & 255)]
    .map(v => v.toString(16).padStart(2, '0')).join('');
}

function reasonText(reason) {
  const messages = {
    border: 'You hit the border.',
    wall: 'You hit a wall.',
    self: 'You ran into yourself.',
    time: 'Time limit reached.',
    full: 'The board is full, you win!',
    stopped: 'Game stopped.'
  };
  return messages[reason] || 'Game over.';
}

function togglePause() {
  const gameState = loadGameState();
  if (!gameState || !gameState.running) return;

  const props = PropertiesService.getDocumentProperties();
  const paused = props.getProperty('snakePaused') === 'true';

  props.setProperty('snakePaused', String(!paused));
  sheet.getRange('H4').setValue(paused ? 'START' : 'PAUSED');
}

function spawnPower(gameState) {
  const taken = new Set();
  gameState.snake.forEach(s => taken.add(s.r + ',' + s.c));
  (gameState.walls || []).forEach(w => taken.add(w.r + ',' + w.c));
  if (gameState.food) taken.add(gameState.food.r + ',' + gameState.food.c);

  const free = [];
  for (let r = 1; r <= size; r++) {
    for (let c = 1; c <= size; c++) {
      if (!taken.has(r + ',' + c)) free.push({ r, c });
    }
  }
  if (free.length === 0) return null;

  const spot = free[Math.floor(Math.random() * free.length)];
  return {
    r: spot.r,
    c: spot.c,
    type: Math.random() < 0.5 ? 'slow' : 'shrink',
    ttl: 40
  };
}

function addLevelWalls(gameState) {
  const head = gameState.snake[0];

  const taken = new Set();
  gameState.snake.forEach(s => taken.add(s.r + ',' + s.c));
  (gameState.walls || []).forEach(w => taken.add(w.r + ',' + w.c));
  if (gameState.food) taken.add(gameState.food.r + ',' + gameState.food.c);
  if (gameState.power) taken.add(gameState.power.r + ',' + gameState.power.c);

  for (let attempt = 0; attempt < 60; attempt++) {
    const horizontal = Math.random() < 0.5;
    const r = Math.floor(Math.random() * size) + 1;
    const c = Math.floor(Math.random() * size) + 1;

    const cells = [0, 1, 2].map(i => ({
      r: horizontal ? r : r + i,
      c: horizontal ? c + i : c
    }));

    const ok = cells.every(cell =>
      cell.r >= 1 && cell.r <= size &&
      cell.c >= 1 && cell.c <= size &&
      !taken.has(cell.r + ',' + cell.c) &&
      Math.abs(cell.r - head.r) + Math.abs(cell.c - head.c) >= 4
    );

    if (ok) {
      cells.forEach(cell => gameState.walls.push(cell));
      return;
    }
  }
}

function setupModeCell() {
  const cell = sheet.getRange('V9:W9');
  const first = sheet.getRange('V9');

  if (!first.isPartOfMerge()) cell.merge();

  if (!first.getDataValidation()) {
    first.setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Borders: deadly', 'Borders: Wrap-around'], true)
        .build()
    );
  }

  first
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(10)
    .setFontWeight('bold')
    .setBackground(lighten(getThemeColor()))
    .setFontColor('#111827');

  if (first.getValue() === '') first.setValue('Borders: deadly');
}

function isWrapMode() {
  return sheet.getRange('V9').getValue() === 'Borders: Wrap-around';
}

function onEdit(e) {
  if (!e || e.range.getSheet().getName() !== board) return;

  if (e.range.getA1Notation() === 'N4') {
    updateHighscores();

    const gameState = loadGameState();
    if (!gameState || !gameState.running) {
      sheet.getRange('A4').setValue('SCORE: 0 | BEST: ' + getBestScore());
    }
  }
}
