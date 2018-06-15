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
    os.system('TASKKILL /F /IM "Minesweeper X.exe"')


def l_click(x, y):
    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)


def r_click(x, y):
    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, x, y, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, x, y, 0, 0)


def screenshotBoard():
    win32api.SetCursorPos((0, 0))
    ig.grab().save("Screenshot.png", "PNG")
    im = image.open('Screenshot.png')
    minesweeper_board = im.crop(box)
    minesweeper_board.save("Board.png")
    os.remove("Screenshot.png")


def read_info_from_board():
    template = cv2.imread("minesweeper_cell.bmp")
    board = cv2.imread("Board.png")
    w, h = template.shape[:-1]
    result = cv2.matchTemplate(board, template, cv2.TM_CCOEFF_NORMED)
    threshold = .8
    loc = np.where(result > threshold)
    print(loc)
    for pt in zip(*loc[::-1]):
        cv2.rectangle(board, pt, (pt[0] + w, pt[1] + h), (0, 0, 255), 2)

    cv2.imwrite("result.png", board)

    for file in os.listdir('digits'):
        print(file)
        small_image = cv2.imread('digits/' + file)
        large_image = cv2.imread("Board.png")
        w, h = small_image.shape[:-1]
        result = cv2.matchTemplate(large_image, small_image, cv2.TM_CCOEFF_NORMED)
        threshold = .8
        loc = np.where(result > threshold)
        print(loc)
        for pt in zip(*loc[::-1]):
            cv2.rectangle(large_image, pt, (pt[0] + w, pt[1] + h), (0, 0, 255), 2)

        cv2.imwrite("result" + file[12] + ".png", large_image)




def main():
    subprocess.Popen(r"MinesweeperX__1.15\Minesweeper X.exe")
    time.sleep(1)
    win32gui.EnumWindows(callback, None)
    screenshotBoard()
    read_info_from_board()
    l_click(box[0] + (box[2] - box[0])//2, box[1] + (box[3]-box[1])//2)
    screenshotBoard()
    read_info_from_board()



if __name__ == '__main__':
    atexit.register(cleanup)
    main()