#!
# -*- coding: utf-8 -*-

"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020-25 https://prrvchr.github.io                                  ║
║                                                                                    ║
║   Permission is hereby granted, free of charge, to any person obtaining            ║
║   a copy of this software and associated documentation files (the "Software"),     ║
║   to deal in the Software without restriction, including without limitation        ║
║   the rights to use, copy, modify, merge, publish, distribute, sublicense,         ║
║   and/or sell copies of the Software, and to permit persons to whom the Software   ║
║   is furnished to do so, subject to the following conditions:                      ║
║                                                                                    ║
║   The above copyright notice and this permission notice shall be included in       ║
║   all copies or substantial portions of the Software.                              ║
║                                                                                    ║
║   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,                  ║
║   EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES                  ║
║   OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.        ║
║   IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY             ║
║   CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,             ║
║   TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE       ║
║   OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.                                    ║
║                                                                                    ║
╚════════════════════════════════════════════════════════════════════════════════════╝
"""

import unohelper

from com.sun.star.awt.PosSize import POS
from com.sun.star.awt.PosSize import POSSIZE

from com.sun.star.awt import Point
from com.sun.star.awt import Rectangle
from com.sun.star.awt import XWindowListener


class Layout(unohelper.Base,
             XWindowListener):
    def __init__(self, window, controls=(), columns=(), links=(), horizontals=(), verticals=(), tops=(), lefts=(), bottoms=()):
        self.window = window
        controls = list(controls)
        columns = list(columns)
        links = list(links)
        horizontals = list(horizontals)
        verticals = list(verticals)
        tops = list(tops)
        lefts = list(lefts)
        bottoms = list(bottoms)
        self._size = window.getPosSize()

        self._controls = []
        for ctrl in controls:
            possize = ctrl.getPosSize()
            right = self._size.Width - (possize.X + possize.Width)
            bottom = self._size.Height - (possize.Y + possize.Height)
            self._controls.append((ctrl, Rectangle(possize.X, possize.Y, right, bottom)))

        self._columns = []
        for col1, col2, grow_y in columns:
            pos1 = col1.getPosSize()
            pos2 = col2.getPosSize()
            bottom1 = self._size.Height - (pos1.Y + pos1.Height)
            bottom2 = self._size.Height - (pos2.Y + pos2.Height)
            self._columns.append((
                col1, pos1.X, pos1.Width, pos1.Y, pos1.Height, bottom1,
                col2, pos2.X, pos2.Width, pos2.Y, pos2.Height, bottom2,
                grow_y
            ))

        self._links = []
        for ctrl, parent in links:
            pos = ctrl.getPosSize()
            parent_pos = parent.getPosSize()
            parent_right = parent_pos.X + parent_pos.Width
            offset_x = pos.X - parent_right
            self._links.append((ctrl, parent, offset_x, pos.Y, pos.Width, pos.Height))

        self._horizontals = []
        for ctrl in horizontals:
            pos = ctrl.getPosSize()
            right = self._size.Width - (pos.X + pos.Width)
            self._horizontals.append((ctrl, Rectangle(pos.X, pos.Y, right, pos.Height)))

        self._verticals = []
        for ctrl in verticals:
            pos = ctrl.getPosSize()
            bottom = self._size.Height - (pos.Y + pos.Height)
            self._verticals.append((ctrl, Rectangle(pos.X, pos.Y, pos.Width, bottom)))

        self._tops = []
        for ctrl in tops:
            pos = ctrl.getPosSize()
            self._tops.append((ctrl, Point(self._size.Width - pos.X, pos.Y)))

        self._lefts = []
        for ctrl in lefts:
            pos = ctrl.getPosSize()
            self._lefts.append((ctrl, Point(pos.X, self._size.Height - pos.Y)))

        self._bottoms = []
        for ctrl in bottoms:
            pos = ctrl.getPosSize()
            self._bottoms.append((ctrl, Point(self._size.Width - pos.X, self._size.Height - pos.Y)))

        self._statics = []
        flat_columns = []
        for item in columns:
            flat_columns.extend([item[0], item[1]])

        handled = set(controls + flat_columns + horizontals + verticals + tops + lefts + bottoms + [item[0] for item in links])
        for ctrl in window.getControls():
            if ctrl not in handled:
                pos = ctrl.getPosSize()
                self._statics.append((ctrl, Point(pos.X, pos.Y)))

    def windowResized(self, event):
        size = event.Source.getPosSize()
        W = size.Width
        H = size.Height

        for ctrl, point in self._statics:
            ctrl.setPosSize(point.X, point.Y, 0, 0, POS)

        for ctrl, possize in self._horizontals:
            width = W - possize.X - possize.Width
            if width > 0:
                ctrl.setPosSize(possize.X, possize.Y, width, possize.Height, POSSIZE)

        for ctrl, possize in self._verticals:
            height = H - possize.Y - possize.Height
            if height > 0:
                ctrl.setPosSize(possize.X, possize.Y, possize.Width, height, POSSIZE)

        for ctrl, point in self._tops:
            ctrl.setPosSize(W - point.X, point.Y, 0, 0, POS)

        for ctrl, point in self._lefts:
            ctrl.setPosSize(point.X, H - point.Y, 0, 0, POS)

        for ctrl, point in self._bottoms:
            ctrl.setPosSize(W - point.X, H - point.Y, 0, 0, POS)

        updated_parents = {}

        if self._columns:
            delta_w = W - self._size.Width
            half_delta = delta_w // 2

            for c1, x1, w1, y1, h1, b1, c2, x2, w2, y2, h2, b2, grow_y in self._columns:
                new_x1 = x1
                new_w1 = w1 + half_delta
                new_x2 = x2 + half_delta
                new_w2 = w2 + half_delta
                new_h1 = H - y1 - b1 if grow_y else h1
                new_h2 = H - y2 - b2 if grow_y else h2

                if new_w1 > 0 and new_h1 > 0:
                    c1.setPosSize(new_x1, y1, new_w1, new_h1, POSSIZE)
                    updated_parents[c1] = (new_x1, new_w1)

                if new_w2 > 0 and new_h2 > 0:
                    c2.setPosSize(new_x2, y2, new_w2, new_h2, POSSIZE)
                    updated_parents[c2] = (new_x2, new_w2)

        for ctrl, parent, x, y, width, height in self._links:
            if parent in updated_parents:
                parent_x, parent_width = updated_parents[parent]
            else:
                possize = parent.getPosSize()
                parent_x, parent_width = possize.X, possize.Width

            ctrl.setPosSize(parent_x + parent_width + x, y, width, height, POSSIZE)

        for ctrl, possize in self._controls:
            width = W - possize.X - possize.Width
            height = H - possize.Y - possize.Height
            if width > 0 and height > 0:
                ctrl.setPosSize(possize.X, possize.Y, width, height, POSSIZE)

    def windowMoved(self, event):
        pass

    def windowShown(self, event):
        pass

    def windowHidden(self, event):
        pass

    def disposing(self, event):
        pass
