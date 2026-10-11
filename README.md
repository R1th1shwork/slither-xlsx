# slither.xlsx

a retro arcade snake engine built completely inside google sheets using apps script for the hack club wrong tool challenge

[play the game here!](https://docs.google.com/spreadsheets/d/1dK1Xdu9ro0hm0D6Q9rLGLb7-uyHXrGUlSg7oGI25rDA/edit?usp=sharing)

## why this exists

google sheets is objectively the absolute worst place on planet earth to build a video game network calls lag screen tearing is real and ui latency is completely cooked we built slither.xlsx anyway just to see how far we could torture a spreadsheet grid until it ran a fully functional snake game plus i am doing all of this because i really want that anbernic rg35xx pro handheld hehe i never had one of the handhelds before in my life so yeeee fingers crossed

## how we fixed the lag so it doesnt melt your cpu

- batch matrix rendering: instead of writing cells one by one like a caveman we build a full 20x20 matrix in memory and dump it all at once using `setValues()`
- forced render flushes: `SpreadsheetApp.flush()` forces synchronous frame execution so the grid actually updates instead of freezing halfway through
- properties service storage: all game state like snake coordinates scores and wall maps live in `PropertiesService` so we dont slow down the game loop reading spreadsheet cells every frame

## features loaded into this monster

- 20x20 pixel grid with 10 goated aesthetic themes ranging from gameboy olive and cyberpunk neon to synthwave dracula matrix and oled black
- universal static power-up colors: even when you cycle through all 10 themes using the dynamic theme switcher the power-ups keep their exact same iconic neon shades so u never get confused
- custom obstacle maps: pick between classic open space box mode or a full labyrinth maze
- 6 insane power-ups:
  - **blue pill**: speed and slow-mo modifier
  - **ghost apple**: gives you 5 seconds of tail phasing immunity so u can clip right through urself without dying
  - **magnet berry**: pulls apples straight toward ur head from a distance
  - **shield rune**: absorbs a brutal crash before u die
  - **shrink tonic**: cuts ur tail length down by 30 percent when things get too tight
  - **diamond gem**: triggers a 3x point multiplier
- arcade death fx: snake blinks white twice and flashes an animation before wiping the board clean
- audio bridge: we tried hooking up web audio api oscillators through an html sidebar to get 8-bit sound triggers going but browser audio permissions proved to be absolute menaces and honestly it doesnt work no matter what we tried i am tired and i gave up on the sound engine lol

## gameplay & how to play guide

1. open up the sheet and hit **RESET** to wipe the canvas and sync properties
2. use the on-sheet directional buttons (`UP`, `DOWN`, `LEFT`, `RIGHT`) to steer your snake
3. pick ur map (`CLASSIC`, `BOX`, `MAZE`) and click the **THEME** button to cycle through all 10 colorways on the fly

**practice mode note (not recommended) 💀**: hitting **START AUTO** runs an automated simulation loop for testing but actually playing it manually with the button controls inside a spreadsheet is pure psychological torture. but hey thats the whole point of the wrong tool challenge
