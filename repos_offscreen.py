import win32gui
import sys
import ctypes
import win32api


SM_CXFULLSCREEN = 16
SM_CYFULLSCREEN = 17


def isAdmin():
    return ctypes.windll.shell32.IsUserAnAdmin() != 0


class Relocator(object):
    def __init__(self, minPixels = 5):
        self.minPixels = 5
        self.desktopWidth = win32api.GetSystemMetrics(SM_CXFULLSCREEN)
        self.desktopHeight = win32api.GetSystemMetrics(SM_CYFULLSCREEN)
        print("Relocator for (primary monitor) desktop size %d x %d" % (self.desktopWidth, self.desktopHeight))
        self.resetStats()
    
    def resetStats(self):
        self.numRelocatedWindows = 0
        self.errors = []

    def relocateAllOffScreenWindows(self):
        print("Relocating all off-screen windows ...")
        self.resetStats()
        win32gui.EnumWindows(self._processWindow, None)
        print("\n=======================\n\nNumber of windows moved: %d" % self.numRelocatedWindows)
        if len(self.errors) > 0:
            print("Number of errors: %d\n  %s" % (len(self.errors), "\n  ".join(self.errors)))
            if not isAdmin():
                print("If errors are due to rights, it may help to run this program as an administrator")

    def isOffPrimaryMonitor(self, hwnd):
        left, top, right, bottom = rect = win32gui.GetWindowRect(hwnd)
        if left == -32000 and top == -32000: # ignore minimised/tray
            isOff = False
        else:
            isOff = left < 0 or top < 0 or left > self.desktopWidth-self.minPixels or top > self.desktopHeight-self.minPixels
        return isOff, rect
        
    def _processWindow(self, hwnd, _):
        isOff, rect = self.isOffPrimaryMonitor(hwnd)
        if isOff:
            windowText = win32gui.GetWindowText(hwnd)
            print("\nWindow 0x%X ('%s') is off-screen with rect=%s" % (hwnd, windowText, rect))
            self.relocateWindow(hwnd, currentRect=rect)

    def relocateWindow(self, hwnd, currentRect=None):
        if currentRect is None:
            currentRect = win32gui.GetWindowRect(hwnd)
        left, top, right, bottom = currentRect
        width = right - left
        height = bottom - top
        try:
            win32gui.MoveWindow(hwnd, 0, 0, width, height, True)
            print("Moved 0x%X from (%d, %d) to (0, 0)" % (hwnd, top, left))
            self.numRelocatedWindows += 1
        except Exception as e:
            errorText = "Error moving 0x%X ('%s'): %s" % (hwnd, win32gui.GetWindowText(hwnd), str(e))
            self.errors.append(errorText)
            print(errorText)


if __name__ == '__main__':
    Relocator().relocateAllOffScreenWindows()
    #Relocator().relocateWindow(0x808d0)
