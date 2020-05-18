#!
# -*- coding: utf_8 -*-

import uno
import unohelper


def getRowSetOrders(rowset):
    orders = [order.strip('"') for order in rowset.Order.replace(' ', '').split(',')]
    return orders

# TODO: On OpenOffice listbox.getSelectedItems() return a "ByteSequence instance" on empty selection
def getSelectedItems(control):
    items = control.getSelectedItems() if control.getSelectedItemPos() != -1 else ()
    return items

# TODO: On OpenOffice listbox.Model.StringItemList return a "ByteSequence instance" on empty listbox
def getStringItemList(control):
    items = control.Model.StringItemList if control.ItemCount != 0 else ()
    return items
