/**
 * SLITHER-XLSX COMPLETE ENGINE WITH DEATH ANIMATION & POWER-UPS
 * 10 Themes + 6 Universal Power-Ups + White Blink Death FX + Control Handlers
 */

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
  ]
};

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
  sheet.getRange("U1:Z20").clearContent();

  props.setProperty("SNAKE", JSON.stringify([[10, 10], [10, 9], [10, 8]]));
  props.setProperty("DIR", "RIGHT");
  props.setProperty("APPLE", JSON.stringify([5, 15]));
  props.setProperty("APPLE_TYPE", "RED");
  props.setProperty("POWERUP", ""); 
  props.setProperty("POWERUP_POS", "[]");
  props.setProperty("GHOST_TIMER", "0");
  props.setProperty("SHIELD_ACTIVE", "false");
  props.setProperty("MULTIPLIER_TIMER", "0");
  props.setProperty("GOLD_TIMER", "0");
  props.setProperty("SCORE", "0");
  props.setProperty("PENDING_SOUND", "");

  if (!props.getProperty("MAP")) props.setProperty("MAP", "CLASSIC");
  if (!props.getProperty("HIGH_SCORE")) props.setProperty("HIGH_SCORE", "0");
  if (!props.getProperty("THEME")) props.setProperty("THEME", "GAMEBOY");

  let headerRange = sheet.getRange("V1:W1");
  headerRange.merge();
  headerRange.setValue("SLITHER HD");
  headerRange.setBackground("#000000");
  headerRange.setFontColor("#00FF66");
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");

  sheet.getRange("V2").setValue("SCORE:");
  sheet.getRange("V3").setValue("BEST:");
  sheet.getRange("V4").setValue("MAP:");
  sheet.getRange("V5").setValue("THEME:");
  sheet.getRange("V2:V5").setFontWeight("bold").setFontColor("#000000");

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

    if (Math.random() < 0.20 && !activePowerup) {
      let pTypes = ["BLUE_PILL", "GHOST", "MAGNET", "SHIELD", "SHRINK", "DIAMOND"];
      let chosenP = pTypes[Math.floor(Math.random() * pTypes.length)];
      let pRow = Math.floor(Math.random() * GRID_HEIGHT);
      let pCol = Math.floor(Math.random() * GRID_WIDTH);
      props.setProperty("POWERUP", chosenP);
      props.setProperty("POWERUP_POS", JSON.stringify([pRow, pCol]));
    }

    let isGold = Math.random() < 0.25;
    appleType = isGold ? "GOLD" : "RED";
    props.setProperty("APPLE_TYPE", appleType);

    let newRow, newCol, invalid;
    do {
      invalid = false;
      newRow = Math.floor(Math.random() * GRID_HEIGHT);
      newCol = Math.floor(Math.random() * GRID_WIDTH);
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

function setMapClassic() { setMap("CLASSIC"); }
function setMapBox()     { setMap("BOX"); }
function setMapMaze()    { setMap("MAZE"); }

function setMap(mapName) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("MAP", mapName);
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
  sheet.getRange("U1:U20").clearContent();
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
