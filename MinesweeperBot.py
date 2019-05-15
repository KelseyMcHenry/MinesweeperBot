import subprocess
import time
import PIL.Image as image
import PIL.ImageGrab as ig
import atexit
from itertools import combinations
import math

import win32gui
import win32api
import win32con
import os
import numpy as np
import random

import cv2

box = ()
coordinates = list()
board_model = list()
coord_board_map = dict()
board_coord_map = dict()
cell_size = 0
X = 0
Y = 0
move_queue = list()
debug_move_queue = list()
done = False
large_number_value = 2
iteration = 0


"""Gets location information about Minesweeper"""
def callback(hwnd, extra):
    rect = win32gui.GetWindowRect(hwnd)
    x = rect[0]
    y = rect[1]
    w = rect[2] - x
    h = rect[3] - y
    if win32gui.GetWindowText(hwnd).count('Minesweeper X') > 0 and w > 0 and h > 0:
        print(win32gui.GetWindowText(hwnd))
        print("\tLocation: (%d, %d)" % (x, y))
        print("\t    Size: (%d, %d)" % (w, h))
        global box
        box = (x, y, w + x, h + y)


def cleanup():
    pass
    # os.system('TASKKILL /F /IM "Minesweeper X.exe"')


"""given a tuple of length 2, will left click at the given position in pixels"""
def l_click(pos):
    y = pos[1] + box[1]
    x = pos[0] + box[0]
    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)


"""given a tuple of length 2, will right click at the given position in pixels"""
def r_click(pos):
    y = pos[1] + box[1]
    x = pos[0] + box[0]
    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, x, y, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, x, y, 0, 0)


"""takes a screenshot of the board and crops it, saving it to Board.PNG"""
def screenshot_board():
    # win32api.SetCursorPos((0, 0))
    ig.grab().save("Screenshot.png", "PNG")
    im = image.open('Screenshot.png')
    minesweeper_board = im.crop(box)
    minesweeper_board.save("Board.png")
    # os.remove("Screenshot.png")


"""
initializes the data model of the board state

"""
def init_data_model(loc):
    # convert data from open CV to list of x,y pixel coordinate pairs
    positions = list(zip(loc[1].tolist(), loc[0].tolist()))
    # pickout the first y value
    yval = positions[0][1]
    overall_list = list()
    line_list = list()
    # split out data into rows
    for i in positions:
        if i[1] == yval:
            line_list.append((i[0], i[1]))
        else:
            temp_list = list()
            temp_list.extend(line_list)
            overall_list.append(temp_list)
            line_list.clear()
            yval = i[1]
            line_list.append((i[0], i[1]))
    overall_list.append(line_list)
    print(overall_list)
    global coordinates
    coordinates = overall_list
    global board_model
    board_model = [["_" for _ in row] for row in overall_list]
    global coord_board_map
    for i in range(0, len(coordinates[0])):
        for j in range(0, len(coordinates)):
            coord_board_map[(coordinates[j][i][0], coordinates[j][i][1])] = (j, i)
            board_coord_map[(j, i)] = coordinates[j][i]
    global X
    global Y
    X = len(coordinates[0])
    Y = len(coordinates)
    print("Width in Mines: " + str(X))
    print("Height in Mines : " + str(Y))


def read_info_from_board():
    for i, row in enumerate(board_model):
        for j, cell in enumerate(row):
            if board_model[i][j] == '?':
                board_model[i][j] = '_'

    screenshot_board()
    template = cv2.imread("minesweeper_blank_tile.png")
    board = cv2.imread("Board.png")
    # determine cell size
    w, h = template.shape[:-1]
    global cell_size
    cell_size = (w, h)
    result = cv2.matchTemplate(board, template, cv2.TM_CCOEFF_NORMED)
    threshold = .95
    loc = np.where(result > threshold)
    for pt in zip(*loc[::-1]):
        cv2.rectangle(board, pt, (pt[0] + w, pt[1] + h), (0, 0, 255), 2)

    cv2.imwrite("result.png", board)
    # init board model with data from openCV if there is no model
    if len(board_model) == 0:
        init_data_model(loc)

    for pt in zip(*loc[::-1]):
        board_model[coord_board_map[(pt[0], pt[1])][0]][coord_board_map[(pt[0], pt[1])][1]] = '?'
        cv2.rectangle(board, pt, (pt[0] + w, pt[1] + h), (0, 0, 255), 2)

    cv2.imwrite("result0.png", board)

    for file in os.listdir('digits'):
        if int(file[12]) < large_number_value:
            print(f"Recognizing: {file[12]} ")
            small_image = cv2.imread('digits/' + file)
            large_image = cv2.imread("Board.png")
            w, h = small_image.shape[:-1]
            result = cv2.matchTemplate(large_image, small_image, cv2.TM_CCOEFF_NORMED)
            threshold = .8
            loc = np.where(result > threshold)
            for pt in zip(*loc[::-1]):
                board_model[coord_board_map[(pt[0]-1, pt[1]-1)][0]][coord_board_map[(pt[0]-1, pt[1]-1)][1]] = int(file[12])
                cv2.rectangle(large_image, pt, (pt[0] + w, pt[1] + h), (0, 0, 255), 2)

            cv2.imwrite("result" + file[12] + ".png", large_image)

    template = cv2.imread("minesweeper_flag.png")
    board = cv2.imread("Board.png")
    w, h = template.shape[:-1]
    result = cv2.matchTemplate(board, template, cv2.TM_CCOEFF_NORMED)
    threshold = .8
    loc = np.where(result > threshold)

    for pt in zip(*loc[::-1]):
        board_model[coord_board_map[(pt[0], pt[1])][0]][coord_board_map[(pt[0], pt[1])][1]] = 'flag'
        cv2.rectangle(board, pt, (pt[0] + w, pt[1] + h), (0, 0, 255), 2)

    cv2.imwrite("result_flag.png", board)

    template = cv2.imread("minesweeper_win.png")
    board = cv2.imread("Board.png")
    w, h = template.shape[:-1]
    result = cv2.matchTemplate(board, template, cv2.TM_CCOEFF_NORMED)
    threshold = .95
    loc = np.where(result > threshold)

    global done
    for pt in zip(*loc[::-1]):
        done = True
        cv2.rectangle(board, pt, (pt[0] + w, pt[1] + h), (0, 0, 255), 2)

    cv2.imwrite("result_win.png", board)

    template = cv2.imread("minesweeper_lose.png")
    board = cv2.imread("Board.png")
    w, h = template.shape[:-1]
    result = cv2.matchTemplate(board, template, cv2.TM_CCOEFF_NORMED)
    threshold = .95
    loc = np.where(result > threshold)

    for pt in zip(*loc[::-1]):
        done = True
        cv2.rectangle(board, pt, (pt[0] + w, pt[1] + h), (0, 0, 255), 2)

    cv2.imwrite("result_lose.png", board)


def cell_location(x, y):
    global cell_size
    corner_coord = board_coord_map[(x, y)]

    return corner_coord[0] + (cell_size[0] // 2), corner_coord[1] + (cell_size[1] // 2)


def add_to_move_queue():
    for i, row in enumerate(board_model):
        for j, cell in enumerate(row):
            if type(cell) is int:
                adj = adjacent_to_cell(i, j)
                if count(adj, '?') + count(adj, 'flag') == board_model[i][j]:
                    for entry in adj:
                        if entry[2] == '?':
                            pos = board_coord_map[(entry[0], entry[1])]
                            # print(pos)
                            if (pos[0] + cell_size[0] // 2, pos[1] + cell_size[1] // 2, r_click) not in move_queue:
                                move_queue.append((pos[0] + cell_size[0] // 2, pos[1] + cell_size[1] // 2, r_click))
                                debug_move_queue.append((entry[0], entry[1], r_click))
                elif count(adj, 'flag') == board_model[i][j] and count(adj, '?') > 0:
                    for entry in adj:
                        if entry[2] == '?':
                            pos = board_coord_map[(entry[0], entry[1])]
                            # print(pos)
                            if (pos[0] + cell_size[0] // 2, pos[1] + cell_size[1] // 2, l_click) not in move_queue:
                                move_queue.append((pos[0] + cell_size[0] // 2, pos[1] + cell_size[1] // 2, l_click))
                                debug_move_queue.append((entry[0], entry[1], r_click))
                elif board_model[i][j] - count(adj, 'flag') == 1 and count(adj, '?') == 3 and wall_of_blanks_direction(adj):
                    direction = wall_of_blanks_direction(adj)
                    neighbors = wall_of_blanks_neighbors(adj, direction)
                    for neighbor in neighbors:
                        if isinstance(neighbor[2], int):
                            adj_neighbor = adjacent_to_cell(neighbor[0], neighbor[1])
                            if count(adj_neighbor, '?') == 3 and wall_of_blanks_direction(adj_neighbor) == direction:
                                if board_model[neighbor[0]][neighbor[1]] - count(adj_neighbor, 'flag') == 2:
                                    direction_prime = (i - neighbor[0], j - neighbor[1])
                                    l_click_target = (i + direction[0] + direction_prime[0], j + direction[1] + direction_prime[1])
                                    l_click_location = board_coord_map[l_click_target[0], l_click_target[1]]
                                    current_target = (board_coord_map[(i, j)][0] + box[0] + cell_size[0] // 2, board_coord_map[(i, j)][1] + box[1] + cell_size[1] // 2)
                                    win32api.SetCursorPos(current_target)
                                    current_target = (board_coord_map[(neighbor[0], neighbor[1])][0] + box[0] + cell_size[0] // 2,
                                                      board_coord_map[(neighbor[0], neighbor[1])][1] + box[1] + cell_size[1] // 2)
                                    win32api.SetCursorPos(current_target)
                                    mouse_target = (l_click_location[0] + box[0] + cell_size[0] // 2, l_click_location[1] + box[1] + cell_size[1] // 2)
                                    win32api.SetCursorPos(mouse_target)
                                    move_queue.append((l_click_location[0] + cell_size[0] // 2, l_click_location[1] + cell_size[1] // 2, l_click))
                                    r_click_target = (neighbor[0] + direction[0] - direction_prime[0], neighbor[1] + direction[1] - direction_prime[1])
                                    r_click_location = board_coord_map[r_click_target[0], r_click_target[1]]
                                    mouse_target = (r_click_location[0] + box[0] + cell_size[0] // 2, r_click_location[1] + box[1] + cell_size[1] // 2)
                                    win32api.SetCursorPos(mouse_target)
                                    move_queue.append((r_click_location[0] + cell_size[0] // 2, r_click_location[1] + cell_size[1] // 2, r_click))
                                    print(f"2-1 target added: {l_click_target} based on {i, j, cell} and {neighbor}")

def count(adj, value):
    count = 0
    for entry in adj:
        for number in entry:
            if number == value:
                count += 1
    return count


def adjacent_to_cell(x, y):
    cells = list()
    for i in range(x - 1, x + 2):
        for j in range(y - 1, y + 2):
            if (not (i == x and j == y)) and (i >= 0) and (i < Y) and (j < X) and (j >= 0):
                try:
                    cells.append((i, j, board_model[i][j]))
                except IndexError:
                    pass
    return cells


def wall_of_blanks_direction(adj):
    if len(adj) < 8:
        return None
    combos = [(adj[0], adj[1], adj[2]), (adj[5], adj[6], adj[7]), (adj[0], adj[3], adj[5]), (adj[2], adj[4], adj[7])]
    directions = {0: (-1, 0), 1: (1, 0), 2: (0, -1), 3: (0, 1)}
    for combo in combos:
        if all([data[2] == '?' for data in combo]):
            return directions[combos.index(combo)]
    return None

def wall_of_blanks_neighbors(adj, direction):
    if direction == (1, 0) or direction == (-1, 0):
        return (adj[3], adj[4])
    elif direction == (0, 1) or direction == (0, -1):
        return (adj[1], adj[6])
    else:
        raise Exception(f"Invalid direction: {direction}")


def process_move_queue():
    global iteration
    global done
    global large_number_value
    iteration += 1
    if len(move_queue) == 0 and len(board_model) > 0 and not done:
        if iteration < 20:
            print("No more definitive moves found, switching to random Guessing")
            get_random_blank()
        elif large_number_value < 8:
            large_number_value += 1
        else:
            done = True
    else:
        print(f"Processing Move Queue: Queue size = {len(move_queue)}")
        for move in move_queue:
            move[2]((move[0], move[1]))
            time.sleep(.02)
        move_queue.clear()


def get_random_blank():
    selected = False
    while not selected:
        rX = random.randint(0, X-1)
        rY = random.randint(0, Y-1)
        if board_model[rY][rX] == '?':
            pos = board_coord_map[(rY, rX)]
            # print(pos)
            if (pos[0] + cell_size[0] // 2, pos[1] + cell_size[1] // 2, l_click) not in move_queue:
                print(f"Random Blank selected: {(rX, rY)}")
                move_queue.append((pos[0] + cell_size[0] // 2, pos[1] + cell_size[1] // 2, l_click))
                selected = True

def main():
    # kills existing minesweepers
    os.system('TASKKILL /F /IM "Minesweeper X.exe"')
    # opens minesweeper
    subprocess.Popen(r"MinesweeperX__1.15\Minesweeper X.exe")
    # delay to let it load
    time.sleep(.5)
    # look for Minesweeper
    win32gui.EnumWindows(callback, None)

    # gather board size, initialize models
    read_info_from_board()

    #click in the middle
    l_click(cell_location(Y//2 - 1, X//2 - 1))
    time.sleep(.5)

    while not done:
        read_info_from_board()
        add_to_move_queue()
        process_move_queue()



if __name__ == '__main__':
    atexit.register(cleanup)
    main()