const GRID_WIDTH = 20;  // Columns A to T
const GRID_HEIGHT = 20; // Rows 1 to 20

const MAPS = {
  CLASSIC: [],
  BOX: [
    ...Array.from({length: 14}, (_, i) => [3, i + 3]),
    ...Array.from({length: 14}, (_, i) => [16, i + 3]),
    ...Array.from({length: 12}, (_, i) => [i + 4, 3]),
    ...Array.from({length: 12}, (_, i) => [i + 4, 16])
  ],
  MAZE: [
    [2, 5], [2, 6], [2, 7], [2, 8], [2, 11], [2, 12], [2, 13], [2, 14],
    [5, 2], [6, 2], [7, 2], [12, 2], [13, 2], [14, 2],
    [5, 17], [6, 17], [7, 17], [12, 17], [13, 17], [14, 17],
    [17, 5], [17, 6], [17, 7], [17, 8], [17, 11], [17, 12], [17, 13], [17, 14],
    [6, 6], [6, 7], [6, 8], [6, 11], [6, 12], [6, 13],
    [13, 6], [13, 7], [13, 8], [13, 11], [13, 12], [13, 13],
    [9, 9], [9, 10], [10, 9], [10, 10]
  ],
  CROSS: [
    ...Array.from({length: 12}, (_, i) => [i + 4, 9]),
    ...Array.from({length: 12}, (_, i) => [i + 4, 10]),
    ...Array.from({length: 12}, (_, i) => [9, i + 4]),
    ...Array.from({length: 12}, (_, i) => [10, i + 4])
  ],
  TUNNEL: [
    ...Array.from({length: 16}, (_, i) => [5, i + 2]),
    ...Array.from({length: 16}, (_, i) => [14, i + 2])
  ],
  PILLARS: [
    [4, 4], [4, 5], [5, 4], [5, 5],
    [4, 14], [4, 15], [5, 14], [5, 15],
    [14, 4], [14, 5], [15, 4], [15, 5],
    [14, 14], [14, 15], [15, 14], [15, 15]
  ],
  SPIRAL: [
    ...Array.from({length: 14}, (_, i) => [2, i + 2]),
    ...Array.from({length: 14}, (_, i) => [i + 2, 17]),
    ...Array.from({length: 12}, (_, i) => [17, i + 4]),
    ...Array.from({length: 10}, (_, i) => [i + 5, 4]),
    ...Array.from({length: 8}, (_, i) => [6, i + 6])
  ]
};

const MAP_ORDER = ["CLASSIC", "BOX", "MAZE", "CROSS", "TUNNEL", "PILLARS", "SPIRAL"];

const POWERUP_COLORS = {
  BLUE_PILL: "#007AFF",   // Val 6
  GHOST: "#FFFFFF",       // Val 7
  MAGNET: "#AF52DE",      // Val 8
  SHIELD: "#FF9500",      // Val 9
  SHRINK: "#34C759",      // Val 10
  DIAMOND: "#5AC8FA"      // Val 11
};

const THEMES = {
  GAMEBOY:     { bg: "#8B956D", body: "#2E4600", head: "#1A2B00", apple: "#800000", wall: "#4A5D23", gold: "#D4AF37" },
  CYBERPUNK:   { bg: "#120224", body: "#00F0FF", head: "#FF007F", apple: "#FFD700", wall: "#3B1443", gold: "#00FF66" },
  OLED:        { bg: "#000000", body: "#FFFFFF", head: "#A0A0A0", apple: "#FF3333", wall: "#333333", gold: "#FFCC00" },
  SYNTHWAVE:   { bg: "#2b213a", body: "#ff7edb", head: "#f97068", apple: "#38ef7d", wall: "#44315f", gold: "#ffe600" },
  DRACULA:     { bg: "#282a36", body: "#50fa7b", head: "#8be9fd", apple: "#ff5555", wall: "#44475a", gold: "#f1fa8c" },
  MATRIX:      { bg: "#0d0208", body: "#00ff41", head: "#ffffff", apple: "#008f11", wall: "#003b00", gold: "#adff2f" },
  SOLARIZED:   { bg: "#073642", body: "#2aa198", head: "#268bd2", apple: "#dc322f", wall: "#002b36", gold: "#b58900" },
  TOKYO_NIGHT: { bg: "#1a1b26", body: "#7aa2f7", head: "#bb9af7", apple: "#f7768e", wall: "#24283b", gold: "#e0af68" },
  COFFEE:      { bg: "#2c221e", body: "#d4a373", head: "#faedcd", apple: "#bc6c25", wall: "#3d312a", gold: "#dda15e" },
  RETRO_ARCADE:{ bg: "#1f1f1f", body: "#ffb703", head: "#fb8500", apple: "#d90429", wall: "#333333", gold: "#8ecae6" }
};

const THEME_ORDER = ["GAMEBOY", "CYBERPUNK", "OLED", "SYNTHWAVE", "DRACULA", "MATRIX", "SOLARIZED", "TOKYO_NIGHT", "COFFEE", "RETRO_ARCADE"];

function setupBoard() {
  const props = PropertiesService.getScriptProperties();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  sheet.getRange("A21:Z30").clearContent();
  sheet.getRange("U1:Z25").clearContent();

  if (!props.getProperty("MAP")) props.setProperty("MAP", "CLASSIC");
  let currentMapName = props.getProperty("MAP");
  let obstacles = MAPS[currentMapName] || [];

  // Strict Inside-Box Safe Spawn Scans
  let startR = 10, startC = 8;
  let safeFound = false;
  let minR = (currentMapName === "BOX") ? 5 : 2;
  let maxR = (currentMapName === "BOX") ? 14 : GRID_HEIGHT - 2;
  let minC = (currentMapName === "BOX") ? 6 : 4;
  let maxC = (currentMapName === "BOX") ? 14 : GRID_WIDTH - 2;

  for (let r = minR; r <= maxR; r++) {
    for (let c = minC; c <= maxC; c++) {
      let clash = obstacles.some(ob => 
        (ob[0] === r && ob[1] === c) || 
        (ob[0] === r && ob[1] === c-1) || 
        (ob[0] === r && ob[1] === c-2)
      );
      if (!clash) {
        startR = r;
        startC = c;
        safeFound = true;
        break;
      }
    }
    if (safeFound) break;
  }

  let initialSnake = [[startR, startC], [startR, startC - 1], [startR, startC - 2]];
  props.setProperty("SNAKE", JSON.stringify(initialSnake));
  props.setProperty("DIR", "RIGHT");
  
  // Safe Apple Spawn Constrained to Box Interior if Box Map
  let aRow, aCol, aInvalid;
  do {
    aInvalid = false;
    if (currentMapName === "BOX") {
      aRow = Math.floor(Math.random() * 10) + 5; // rows 5 to 14 inside box
      aCol = Math.floor(Math.random() * 10) + 5; // cols 5 to 14 inside box
    } else {
      aRow = Math.floor(Math.random() * GRID_HEIGHT);
      aCol = Math.floor(Math.random() * GRID_WIDTH);
    }
    for (let ob of obstacles) if (ob[0] === aRow && ob[1] === aCol) aInvalid = true;
    for (let s of initialSnake) if (s[0] === aRow && s[1] === aCol) aInvalid = true;
  } while (aInvalid);

  props.setProperty("APPLE", JSON.stringify([aRow, aCol]));
  props.setProperty("APPLE_TYPE", "RED");
  props.setProperty("POWERUP", ""); 
  props.setProperty("POWERUP_POS", "[]");
  props.setProperty("GHOST_TIMER", "0");
  props.setProperty("SHIELD_ACTIVE", "false");
  props.setProperty("MULTIPLIER_TIMER", "0");
  props.setProperty("GOLD_TIMER", "0");
  props.setProperty("SCORE", "0");
  props.setProperty("PENDING_SOUND", "");

  if (!props.getProperty("HIGH_SCORE")) props.setProperty("HIGH_SCORE", "0");
  if (!props.getProperty("THEME")) props.setProperty("THEME", "GAMEBOY");

  // Arcade Header
  let headerRange = sheet.getRange("V1:W1");
  headerRange.merge();
  headerRange.setValue("SLITHER.XLSX");
  headerRange.setBackground("#000000");
  headerRange.setFontColor("#00FF66");
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");

  // Dashboard Labels
  sheet.getRange("V2").setValue("SCORE:");
  sheet.getRange("V3").setValue("BEST:");
  sheet.getRange("V4").setValue("MAP:");
  sheet.getRange("V5").setValue("THEME:");
  sheet.getRange("V2:V5").setFontWeight("bold").setFontColor("#000000");

  // Repo Link & Guide
  sheet.getRange("V7").setValue("REPO & README:").setFontWeight("bold");
  sheet.getRange("V8").setValue('HYPERLINK("https://github.com/Rithishwork/slither-xlsx", "GitHub Repository")').setFormula('=HYPERLINK("https://github.com/Rithishwork/slither-xlsx", "GitHub Repository")');
  
  sheet.getRange("V10").setValue("POWER-UP GUIDE:").setFontWeight("bold");
  sheet.getRange("V11").setValue("🔵 Blue Pill: Speed Mod");
  sheet.getRange("V12").setValue("⚪ Ghost: Tail Pass");
  sheet.getRange("V13").setValue("💜 Magnet: Pull Apples");
  sheet.getRange("V14").setValue("🧡 Shield: Crash Absorb");
  sheet.getRange("V15").setValue("💚 Shrink: -30% Length");
  sheet.getRange("V16").setValue("💎 Diamond: 3x Points");

  applyTheme(props.getProperty("THEME"));
  updateDashboard(0, parseInt(props.getProperty("HIGH_SCORE")));
  render();
}

function gameTick() {
  const props = PropertiesService.getScriptProperties();
  let snake = JSON.parse(props.getProperty("SNAKE") || "[[10,10],[10,9],[10,8]]");
  let dir = props.getProperty("DIR") || "RIGHT";
  let apple = JSON.parse(props.getProperty("APPLE") || "[5,15]");
  let appleType = props.getProperty("APPLE_TYPE") || "RED";
  let score = parseInt(props.getProperty("SCORE") || "0");
  let highScore = parseInt(props.getProperty("HIGH_SCORE") || "0");
  let currentMapName = props.getProperty("MAP") || "CLASSIC";
  let obstacles = MAPS[currentMapName] || [];

  let ghostTimer = parseInt(props.getProperty("GHOST_TIMER") || "0");
  let shieldActive = props.getProperty("SHIELD_ACTIVE") === "true";
  let multTimer = parseInt(props.getProperty("MULTIPLIER_TIMER") || "0");
  let activePowerup = props.getProperty("POWERUP") || "";
  let powerupPos = JSON.parse(props.getProperty("POWERUP_POS") || "[]");

  if (ghostTimer > 0) props.setProperty("GHOST_TIMER", (ghostTimer - 1).toString());
  if (multTimer > 0) props.setProperty("MULTIPLIER_TIMER", (multTimer - 1).toString());

  let head = [...snake[0]];
  if (dir === "RIGHT") head[1] += 1;
  if (dir === "LEFT")  head[1] -= 1;
  if (dir === "UP")    head[0] -= 1;
  if (dir === "DOWN")  head[0] += 1;

  let hitWall = head[0] < 0 || head[0] >= GRID_HEIGHT || head[1] < 0 || head[1] >= GRID_WIDTH;
  let hitObstacle = false;
  for (let ob of obstacles) {
    if (head[0] === ob[0] && head[1] === ob[1]) hitObstacle = true;
  }

  let hitSelf = false;
  if (ghostTimer <= 0) {
    for (let segment of snake) {
      if (head[0] === segment[0] && head[1] === segment[1]) hitSelf = true;
    }
  }

  if (hitWall || hitObstacle || hitSelf) {
    if (shieldActive) {
      props.setProperty("SHIELD_ACTIVE", "false");
      triggerSound("GOLD");
    } else {
      handleGameOver(snake, apple, appleType, obstacles);
      return;
    }
  } else {
    snake.unshift(head);
  }

  if (activePowerup === "MAGNET" && apple) {
    if (Math.abs(apple[0] - head[0]) <= 3 && Math.abs(apple[1] - head[1]) <= 3) {
      if (apple[0] < head[0]) apple[0]++;
      else if (apple[0] > head[0]) apple[0]--;
      if (apple[1] < head[1]) apple[1]++;
      else if (apple[1] > head[1]) apple[1]--;
      props.setProperty("APPLE", JSON.stringify(apple));
    }
  }

  if (powerupPos.length === 2 && head[0] === powerupPos[0] && head[1] === powerupPos[1]) {
    triggerSound("GOLD");
    if (activePowerup === "GHOST") props.setProperty("GHOST_TIMER", "15");
    if (activePowerup === "SHIELD") props.setProperty("SHIELD_ACTIVE", "true");
    if (activePowerup === "SHRINK") {
      if (snake.length > 3) snake.splice(Math.floor(snake.length * 0.7));
    }
    if (activePowerup === "DIAMOND") props.setProperty("MULTIPLIER_TIMER", "20");

    props.setProperty("POWERUP", "");
    props.setProperty("POWERUP_POS", "[]");
  }

  if (head[0] === apple[0] && head[1] === apple[1]) {
    triggerSound((appleType === "GOLD") ? "GOLD" : "EAT");

    let mult = (multTimer > 0) ? 3 : 1;
    let points = ((appleType === "GOLD") ? 30 : 10) * mult;
    score += points;
    props.setProperty("SCORE", score.toString());

    if (score > highScore) {
      highScore = score;
      props.setProperty("HIGH_SCORE", highScore.toString());
    }

    updateDashboard(score, highScore);

    // Bulletproof Power-Up Spawn (Strictly inside box if box map)
    if (Math.random() < 0.20 && !activePowerup) {
      let pTypes = ["BLUE_PILL", "GHOST", "MAGNET", "SHIELD", "SHRINK", "DIAMOND"];
      let chosenP = pTypes[Math.floor(Math.random() * pTypes.length)];
      let pRow, pCol, pInvalid;
      do {
        pInvalid = false;
        if (currentMapName === "BOX") {
          pRow = Math.floor(Math.random() * 10) + 5;
          pCol = Math.floor(Math.random() * 10) + 5;
        } else {
          pRow = Math.floor(Math.random() * GRID_HEIGHT);
          pCol = Math.floor(Math.random() * GRID_WIDTH);
        }
        for (let ob of obstacles) if (ob[0] === pRow && ob[1] === pCol) pInvalid = true;
        for (let s of snake) if (s[0] === pRow && s[1] === pCol) pInvalid = true;
        if (apple[0] === pRow && apple[1] === pCol) pInvalid = true;
      } while (pInvalid);
      props.setProperty("POWERUP", chosenP);
      props.setProperty("POWERUP_POS", JSON.stringify([pRow, pCol]));
    }

    let isGold = Math.random() < 0.25;
    appleType = isGold ? "GOLD" : "RED";
    props.setProperty("APPLE_TYPE", appleType);

    let newRow, newCol, invalid;
    do {
      invalid = false;
      if (currentMapName === "BOX") {
        newRow = Math.floor(Math.random() * 10) + 5;
        newCol = Math.floor(Math.random() * 10) + 5;
      } else {
        newRow = Math.floor(Math.random() * GRID_HEIGHT);
        newCol = Math.floor(Math.random() * GRID_WIDTH);
      }
      for (let ob of obstacles) if (ob[0] === newRow && ob[1] === newCol) invalid = true;
      for (let s of snake) if (s[0] === newRow && s[1] === newCol) invalid = true;
    } while (invalid);

    apple = [newRow, newCol];
    props.setProperty("APPLE", JSON.stringify(apple));
  } else {
    snake.pop();
  }

  props.setProperty("SNAKE", JSON.stringify(snake));
  renderState(snake, apple, appleType, obstacles, activePowerup, powerupPos, false);
}

function handleGameOver(snake, apple, appleType, obstacles) {
  triggerSound("GAMEOVER");

  for (let b = 0; b < 2; b++) {
    renderState(snake, apple, appleType, obstacles, "", [], true);
    Utilities.sleep(150);
    renderState(snake, apple, appleType, obstacles, "", [], false);
    Utilities.sleep(150);
  }

  setupBoard();
}

function startAutoGame() {
  for (let i = 0; i < 500; i++) {
    gameTick();
    Utilities.sleep(100);
  }
}

function cycleMap() {
  const props = PropertiesService.getScriptProperties();
  const currentMap = props.getProperty("MAP") || "CLASSIC";
  
  let currentIndex = MAP_ORDER.indexOf(currentMap);
  currentIndex = (currentIndex === -1 || currentIndex >= MAP_ORDER.length - 1) ? 0 : currentIndex + 1;
  
  let nextMap = MAP_ORDER[currentIndex];
  props.setProperty("MAP", nextMap);
  setupBoard();
}

function updateDashboard(score, highScore) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const props = PropertiesService.getScriptProperties();

  sheet.getRange("W2").setValue(score).setFontWeight("bold").setHorizontalAlignment("right");
  sheet.getRange("W3").setValue(highScore).setFontWeight("bold").setHorizontalAlignment("right");
  sheet.getRange("W4").setValue(props.getProperty("MAP") || "CLASSIC").setFontWeight("bold").setHorizontalAlignment("right");
  sheet.getRange("W5").setValue(props.getProperty("THEME") || "GAMEBOY").setFontWeight("bold").setHorizontalAlignment("right");
}

function toggleTheme() {
  const props = PropertiesService.getScriptProperties();
  const currentTheme = props.getProperty("THEME") || "GAMEBOY";
  
  let currentIndex = THEME_ORDER.indexOf(currentTheme);
  currentIndex = (currentIndex === -1 || currentIndex >= THEME_ORDER.length - 1) ? 0 : currentIndex + 1;
  
  let nextTheme = THEME_ORDER[currentIndex];
  applyTheme(nextTheme);
  updateDashboard(parseInt(props.getProperty("SCORE") || "0"), parseInt(props.getProperty("HIGH_SCORE") || "0"));
}

function applyTheme(themeKey) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const theme = THEMES[themeKey] || THEMES.GAMEBOY;
  const range = sheet.getRange("A1:T20");

  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0).setBackground(theme.bg).setFontColor(theme.bg).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(1).setBackground(theme.body).setFontColor(theme.body).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(2).setBackground(theme.head).setFontColor(theme.head).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(3).setBackground(theme.apple).setFontColor(theme.apple).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(4).setBackground(theme.wall).setFontColor(theme.wall).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(5).setBackground(theme.gold).setFontColor(theme.gold).setRanges([range]).build(),

    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(6).setBackground(POWERUP_COLORS.BLUE_PILL).setFontColor(POWERUP_COLORS.BLUE_PILL).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(7).setBackground(POWERUP_COLORS.GHOST).setFontColor(POWERUP_COLORS.GHOST).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(8).setBackground(POWERUP_COLORS.MAGNET).setFontColor(POWERUP_COLORS.MAGNET).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(9).setBackground(POWERUP_COLORS.SHIELD).setFontColor(POWERUP_COLORS.SHIELD).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(10).setBackground(POWERUP_COLORS.SHRINK).setFontColor(POWERUP_COLORS.SHRINK).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(11).setBackground(POWERUP_COLORS.DIAMOND).setFontColor(POWERUP_COLORS.DIAMOND).setRanges([range]).build()
  ];

  sheet.setConditionalFormatRules(rules);
  PropertiesService.getScriptProperties().setProperty("THEME", themeKey);
  SpreadsheetApp.flush();
}

function renderState(snake, apple, appleType, obstacles, activePowerup, powerupPos, isWhiteFlash) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let grid = Array(GRID_HEIGHT).fill(0).map(() => Array(GRID_WIDTH).fill(0));

  for (let ob of obstacles) if (ob[0] < GRID_HEIGHT && ob[1] < GRID_WIDTH) grid[ob[0]][ob[1]] = 4;
  if (apple[0] < GRID_HEIGHT && apple[1] < GRID_WIDTH) grid[apple[0]][apple[1]] = (appleType === "GOLD") ? 5 : 3;

  for (let i = 1; i < snake.length; i++) {
    if (snake[i][0] < GRID_HEIGHT && snake[i][1] < GRID_WIDTH) {
      grid[snake[i][0]][snake[i][1]] = isWhiteFlash ? 7 : 1;
    }
  }
  if (snake[0][0] < GRID_HEIGHT && snake[0][1] < GRID_WIDTH) {
    grid[snake[0][0]][snake[0][1]] = isWhiteFlash ? 7 : 2;
  }

  if (activePowerup && powerupPos.length === 2) {
    const pVal = { "BLUE_PILL": 6, "GHOST": 7, "MAGNET": 8, "SHIELD": 9, "SHRINK": 10, "DIAMOND": 11 }[activePowerup] || 6;
    grid[powerupPos[0]][powerupPos[1]] = pVal;
  }

  sheet.getRange(1, 1, GRID_HEIGHT, GRID_WIDTH).setValues(grid);
  sheet.getRange("U1:U20").clearContent();
  SpreadsheetApp.flush();
}

function render() {
  const props = PropertiesService.getScriptProperties();
  let snake = JSON.parse(props.getProperty("SNAKE"));
  let apple = JSON.parse(props.getProperty("APPLE"));
  let appleType = props.getProperty("APPLE_TYPE") || "RED";
  let mapName = props.getProperty("MAP") || "CLASSIC";
  let obstacles = MAPS[mapName] || [];
  let activePowerup = props.getProperty("POWERUP") || "";
  let powerupPos = JSON.parse(props.getProperty("POWERUP_POS") || "[]");
  renderState(snake, apple, appleType, obstacles, activePowerup, powerupPos, false);
}

function goUp()    { setDirection("UP"); }
function goDown()  { setDirection("DOWN"); }
function goLeft()  { setDirection("LEFT"); }
function goRight() { setDirection("RIGHT"); }

function setDirection(newDir) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("DIR", newDir);
  gameTick(); 
}

function openAudioEngine() {
  const html = HtmlService.createHtmlOutputFromFile('SoundApp').setTitle('Slither Audio').setWidth(200);
  SpreadsheetApp.getUi().showSidebar(html);
}

function triggerSound(soundType) {
  PropertiesService.getScriptProperties().setProperty("PENDING_SOUND", soundType);
}

function getPendingSound() {
  const props = PropertiesService.getScriptProperties();
  const sound = props.getProperty("PENDING_SOUND") || "";
  if (sound) props.setProperty("PENDING_SOUND", "");
  return sound;
}
