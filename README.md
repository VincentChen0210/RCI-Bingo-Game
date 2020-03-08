# 'RCI' Bingo Game
Bingo game created on Visual Basic 6 in May-June 2018 (Grade 11)

Features:
- 2 playing cards
- 2 different card sizes (5x5 and 6x6)
- Single player mode and Demo mode (auto-run mode)
- 'House' Challenge in single-player mode
- Highscores list

This program is a bingo - styled game which is played by calling numbers through a button and clicking the corresponding tiles (if available).

This bingo game uses control arrays for cards, timers for "House Challenge" and for the Demo mode, and file reading for Highscores.
The program was created with only functions, in addition to a small usage of a code module.

House scoring triggers when the game has detected that the player has a wining row/column/diagonal, but the player has not called RCIGO (aka BINGO).
This allows the player try to score more points, as the game will continue. However, from this point on the player is playing against the computer:
the longer the game continues and the more cards the player calls, there is chance that the computer/'AI' will detect that the player is trying to ramp up their score.
The chance will continue to climb, until the player decides to end the game by calling "RCIGO" or when the computer catches the player. In the case that the player is caught,
the score that the player would have earned at that time will automatically be saved under the computer's name.

PS: "RCIGO" is named after the author's high school, Riverdale Collegiate Institute.
