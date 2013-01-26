from PIL import Image, ImageGrab, ImageFilter
from PIL.ImageOps import posterize
from string import ascii_uppercase
from time import clock, sleep
from math import cos, sin, atan, sqrt, pi
from threading import Thread
import win32api, win32con, traceback
from random import randrange


#At the moment, just 1 class for the whole bot
#Sometime in the future I need to split it up into it's 3 main parts which are:
#   Reading game board
#   Deciding the best most
#   Moving the mouse

class BejeweledBot:

    MOUSE_SPEED = 800
    POS_FUZZ = 10
    NUM_GEMS = 7
    SAMPLE_SIZE = 5
    ROWS, COLS = 8, 8
    WIDTH, HEIGHT = 458, 458
    ORIGIN_X, ORIGIN_Y = 509, 172
    
    CELLW, CELLH = int(WIDTH / COLS), int(HEIGHT / ROWS)
    GAME_BOX = box = (ORIGIN_X, ORIGIN_Y, ORIGIN_X + WIDTH, ORIGIN_Y + HEIGHT)

    def __init__(self):
        self.pos1 = (0, 0)
        self.pos2 = (0, 0)
        self.dirty = False
        self.torun = False
        self.thread = None
        self.matrix = [[None for i in range(self.COLS)] for j in range(self.ROWS)]

    #capture the board
    def getMatrix(self):
        #grab a screen cap and apply some filters to make it easier to work with
        cap = ImageGrab.grab(self.GAME_BOX)
        cap = cap.filter(ImageFilter.MinFilter)
        cap = cap.filter(ImageFilter.MaxFilter)
        cap = posterize(cap, 3)
        
        #first pass
        #look for color values in the expected gem locations
        #group identical color values together - same gem type
        colors = dict()
        for col in range(self.COLS):
            for row in range(self.ROWS):
                cx = (row * self.CELLW) + self.CELLW / 2
                cy = (col * self.CELLH) + self.CELLH / 2     
                sample_rect = (cx - 3, cy + 14,
                               cx + 3, cy + 21)
                sample_rect = map(int, sample_rect)
                sample = cap.crop(sample_rect)
                color = sorted(sample.getcolors((2 * self.SAMPLE_SIZE) ** 2))[-1][1]
                self.matrix[col][row] = color
                if color in colors:
                    colors[color] += 1
                else:
                    colors[color] = 1
    
        #break if we found too little or too many colors on first pass
        numColors = len(colors)
        if numColors < self.NUM_GEMS or numColors > len(ascii_uppercase):
            return None
        
        #almost always going to be more gems type than there should be
        #merge closest pixel value groups until the expected amount of gems is reached
        #keep track of the ones we merged, probably flashing gems so worth extra points
        color_map = dict(zip(colors, ascii_uppercase))
        for i in range(len(colors) - self.NUM_GEMS): 
            rarest = min(colors, key=colors.get)
            minDist, merge = float("inf"), None
            for color in colors:
                if not color == rarest:
                    #use geometric distance between RGB values (my own idea)
                    dist = sum([(a - b) ** 2 for a, b in zip(color, rarest)])
                    if dist < minDist:
                        minDist = dist
                        merge = color
            
            colors[merge] += 1
            colors[rarest] = colors[merge]
            color_map[rarest] = color_map[merge]
        
        #make matrix easier to read by replacing values with letters
        #mostly for debugging
        for row in self.matrix:
            for index, color in enumerate(row):
                row[index] = color_map[color]
                    
        return self.matrix

    #evaluates the points to be won on a board
    def evalBoard(self):
        total = 0
        for r in range(self.ROWS):
            for c in range(self.COLS):
                gem = self.matrix[r][c]

                #horizontal
                for i in range(1, self.COLS - c):
                    if self.matrix[r][c + i] == gem:
                        if i >= 2:
                            total += 1
                    else:
                        break

                #vertical
                for j in range(1, self.ROWS - r):
                    if self.matrix[r + j][c] == gem:
                        if j >= 2:
                            total += 1
                    else:
                        break
        return total

    #swap two positions in the matrix
    def swapMatrix(self, pos1, pos2):
        self.matrix[pos1[1]][pos1[0]], self.matrix[pos2[1]][pos2[0]] = self.matrix[pos2[1]][pos2[0]], self.matrix[pos1[1]][pos1[0]]

    #save the best swap found
    def rankSwap(self, pos1, pos2):
        #swap positions
        self.swapMatrix(pos1, pos2)
        score = self.evalBoard()
        if score > self.best:
            self.best = score
            self.pair = pos1, pos2
        #swap back to original
        self.swapMatrix(pos1, pos2)    

    #run through possible position swaps
    def findBestSwap(self):
        self.best, self.pair = 0, []
        for row in range(self.ROWS):
            for col in range(self.COLS):
                if row < self.ROWS - 1:            
                    self.rankSwap((col, row), (col, row + 1))
                if col < self.COLS - 1:
                    self.rankSwap((col, row), (col + 1, row))
        if len(self.pair) == 2:
            return self.pair
        else:
            return None

    #given a grid position, return the mouse coordinates of that position
    def gridToMouse(self, pos):
        return(self.ORIGIN_X + pos[0] * self.CELLW + int(self.CELLW / 2),
                self.ORIGIN_Y + pos[1] * self.CELLH + int(self.CELLH / 2))

    #move the mouse is an arc across the board to simulate human inefficiency and disguise the bot
    def moveMouse(self, pos1, pos2):
        if self.dirty: 
            return

        dx = pos2[0] - pos1[0]
        dy = pos2[1] - pos1[1]

        if dx == 0: dx = 1
        if dy == 0: dy = 1

        ang = atan(-dx / dy)
        dist = sqrt(dx * dx + dy * dy)
        xCoeff, yCoeff = cos(ang), sin(ang)

        startTime = clock()
        totalTime = dist / self.MOUSE_SPEED
        while clock() < startTime + totalTime and not self.dirty:
            pc = (clock() - startTime) / totalTime
            arcLen = sin(pi * pc) * (dist / 5)
            x = pos1[0] + (dx * pc) + (xCoeff * arcLen)
            y = pos1[1] + (dy * pc) + (yCoeff * arcLen)
            win32api.SetCursorPos(self.fuzzPos((int(x), int(y)), 2))

    #fuzz a given position, used to further disguise mouse movements
    def fuzzPos(self, pos, fuzz=POS_FUZZ):
        #dont fuzz y down so mouse doesnt get in way of sample
        return (pos[0] + randrange(-fuzz, fuzz), pos[1] - randrange(fuzz))

    #thread controlling the mouse movements
    #will change targets if a new position is found while moving
    def moveThread(self, pos1, pos2):
        self.dirty = False
        pos0 = win32api.GetCursorPos()
        pos1, pos2 = map(self.gridToMouse, [pos1, pos2])
        pos1, pos2 = map(self.fuzzPos, [pos1, pos2])
        self.moveMouse(pos0, pos1)
        self.moveMouse(pos1, pos2)
        self.pos1, self.pos2 = (0, 0), (0, 0)

    #perform the swap using the mouse
    def swap(self, pos1, pos2):
        if not (self.pos1 == pos1 and self.pos2 == pos2):
            self.dirty = True
            self.pos1 = pos1
            self.pos2 = pos2
            if self.thread and self.thread.isAlive():
                self.thread.join()
            self.thread = Thread(target=self.moveThread, args=(pos1, pos2))
            self.thread.start()
        
    #main method
    #stop & stop bot using keyboard to regain mouse control
    #use some basic logging to see whats going on
    def run(self):
        while True:
            if not win32api.GetAsyncKeyState(ord('D')) == 0:
                self.torun = True
            if not win32api.GetAsyncKeyState(ord('A')) == 0:
                self.torun = False
            if self.torun:
                if self.getMatrix():
                    if self.findBestSwap():
                        self.swap(*self.pair)
                        print("Swapping", *self.pair)
                    else:
                        print("No solutions found")
                else:
                    print("Gem count is out of threshold")
        
bot = BejeweledBot()
bot.run()