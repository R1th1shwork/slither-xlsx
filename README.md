# slither.xlsx

a retro arcade snake engine built inside google sheets using apps script for the hack club wrong tool challenge

## why this exists
google sheets is objectively the worst place to build a video game. network calls lag, screen tearing is real, and the ui latency is cooked. we built slither.xlsx anyway to see how far we could push spreadsheet grid rendering.

## how we fixed the lag
* batch matrix rendering: instead of writing individual cells in loops, we construct a full 20x20 matrix in memory and push it all at once using setValues()
* forced render flushes: used SpreadsheetApp.flush() right after updates to force synchronous frame execution so the screen actually updates without tearing
* zero-read state management: all game state (snake position, score, active powerups, wall maps) lives in PropertiesService so we don't do slow spreadsheet reads during game loops

## features
* 20x20 pixel grid with retro game boy styling
* custom obstacle maps (classic, box, and full labyrinth maze)
* golden apple power-ups (+30 pts) with a 15-tick expiration timer
* persistent top 5 leaderboard with timestamps
* continuous auto-run loop (~100ms ticks)

## how to run
1. open the sheet and hit RESET to wipe the canvas and sync properties
2. hit START AUTO to trigger the main loop
3. steer using the directional buttons or pick a map with the selector buttons
