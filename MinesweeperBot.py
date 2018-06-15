import subprocess
import time
import PIL.Image as image
import PIL.ImageGrab as ig
import atexit

import win32gui
import win32api
import win32con
import os
import numpy as np

import cv2

box = ()
coordinates = list()
board_model = list()
coord_board_map = dict()
board_coord_map = dict()
cell_size = 0
X = 0
Y = 0


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


def l_click(pos):
    y = pos[1] + box[1]
    x = pos[0] + box[0]
    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)


def r_click(pos):
    y = pos[1] + box[1]
    x = pos[0] + box[0]
    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, x, y, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, x, y, 0, 0)


def screenshot_board():
    # win32api.SetCursorPos((0, 0))
    ig.grab().save("Screenshot.png", "PNG")
    im = image.open('Screenshot.png')
    minesweeper_board = im.crop(box)
    minesweeper_board.save("Board.png")
    os.remove("Screenshot.png")


def init_data_model(loc):
    x = loc[0]
    y = loc[1]
    positions = list(zip(x.tolist(), y.tolist()))
    xval = positions[0][0]
    overall_list = list()
    line_list = list()
    for i in positions:
        if i[0] == xval:
            line_list.append((i[1], i[0]))
        else:
            temp_list = list()
            temp_list.extend(line_list)
            overall_list.append(temp_list)
            line_list.clear()
            xval = i[0]
            line_list.append((i[1], i[0]))
    overall_list.append(line_list)

    global coordinates
    coordinates = overall_list
    global board_model
    board_model = [["?" for _ in row] for row in overall_list]
    global coord_board_map
    for x in range(0, len(coordinates)):
        for y in range(0, len(coordinates[0])):
            coord_board_map[coordinates[x][y]] = (x, y)
            board_coord_map[(x, y)] = coordinates[x][y]
    global X
    global Y
    X = len(coordinates[0])
    Y = len(coordinates)
    print("X : " + str(X))
    print("Y : " + str(Y))


def read_info_from_board():
    template = cv2.imread("minesweeper_cell.png")
    board = cv2.imread("Board.png")
    w, h = template.shape[:-1]
    global cell_size
    cell_size = (w, h)
    result = cv2.matchTemplate(board, template, cv2.TM_CCOEFF_NORMED)
    threshold = .8
    loc = np.where(result > threshold)
    if len(board_model) == 0:
        init_data_model(loc)

    for pt in zip(*loc[::-1]):
        cv2.rectangle(board, pt, (pt[0] + w, pt[1] + h), (0, 0, 255), 2)

    cv2.imwrite("result.png", board)

    for file in os.listdir('digits'):
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

def cell_location(x, y):
    global cell_size
    corner_coord = board_coord_map[(x, y)]

    return corner_coord[0] + (cell_size[0] // 2), corner_coord[1] + (cell_size[1] // 2)

def main():
    subprocess.Popen(r"MinesweeperX__1.15\Minesweeper X.exe")
    time.sleep(.5)
    win32gui.EnumWindows(callback, None)
    screenshot_board()
    read_info_from_board()
    l_click(cell_location(0, 0))
    # for i in range(0, len(board_model)):
    #     for j in range(0, len(board_model[0])):
    #         r_click(cell_location(i, j))
    # for row in board_model:
    #     print(row)


if __name__ == '__main__':
    atexit.register(cleanup)
    main()