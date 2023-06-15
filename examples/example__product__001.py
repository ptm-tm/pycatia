#! /usr/bin/python3.9

"""

    Example - Product - 001

    Description:
        Assembly Design: Reorder a Product tree alphabetically. The Product shall already be loaded.
        Sort a product tree like this:

            Product1
                |- Product2 (Product2.1)
                |- Product3 (Product3.1)
                |- Product4 (Product4.1)
                |- Part1 (Part1.1)
                |- Part3 (Part3.1)
                |- Part2 (Part2.1)
                |- Product2 (Product2.2)
                |- Product2 (Product2.3)
                |- Part3 (Part3.2)
                |- Part3 (Part3.3)

        To this:

            Product1
                |- Part1 (Part1.1)
                |- Part2 (Part2.1)
                |- Part3 (Part3.1)
                |- Part3 (Part3.2)
                |- Part3 (Part3.3)
                |- Product2 (Product2.1)
                |- Product2 (Product2.2)
                |- Product2 (Product2.3)
                |- Product3 (Product3.1)
                |- Product4 (Product4.1)

    Requirements:
        - pywinauto installed (pip install pywinauto).
        - An open product document with children that need sorting.
        
    Warnings:
        With regards to pycatia this example only shows how to select the root
        product. The rest is handled by pywinauto. _https://pywinauto.github.io/

        You will need to manually install package pywinauto to run this script.
        Also, the placement of `from pywinauto import Desktop` is important.

        This seems to fragile. Sometimes it works for me sometimes it doesn't.

"""

##########################################################
# insert syspath to project folder so examples can be run.
# for development purposes.
import os
import sys

sys.path.insert(0, os.path.abspath("..\\pycatia"))
##########################################################

# todo: explore why this is so fragile.

from pywinauto import Desktop

from pycatia import catia
from pycatia.product_structure_interfaces.product import Product
from pycatia.product_structure_interfaces.product_document import ProductDocument

# Keep in mind, that almost all CATIA commands and windows are dependent on the UI language.
# In order for this example to work you need to set your CATIA to english, or use the correct
# translation of the command or window text.
reorder_cmd_name = "Graph tree reordering"
# reorder_cmd_name = "Neuordnung des Grafikbaums"
reorder_window_name = "Graph tree reordering"
# reorder_window_name = "Neuordnung des Grafikstrukturbaums"
window_text_ok = "OK"
# window_text_ok = "OK"
window_text_apply = "Apply"
# window_text_apply = "Anwenden"
window_text_move_up = "Move Up"
# window_text_move_up = "Nach oben verschieben"
window_text_move_down = "Move Down"
# window_text_move_down = "Nach unten verschieben"
window_text_list_box = "ListProd"
# window_text_list_box = "ListProd"


caa = catia()
document = ProductDocument(caa.active_document.com_object)
product = Product(document.product.com_object)
# Note: It's not necessary to explicitly use the ProductDocument or the Product class
# with the com_object. It's perfectly fine to write it like this:
#   document = caa.active_document
#   product = document.product
# But declaring 'document' and 'product' this way, your linter can't resolve the
# product reference, see https://github.com/evereux/pycatia/issues/107#issuecomment-1336195688

# create the selection object.
selection = document.selection
# add the product to the selection.
selection.add(product)
# open the Graph tree reordering window.
caa.start_command(reorder_cmd_name)
# that's it for pycatia.

# if the import is placed at the top of the file the pycatia steps above will
# not work.

windows = Desktop().windows()

graph_window = None
for window in windows:
    if "tree" in window.window_text():
        print(window.window_text())
    if reorder_window_name in window.window_text():
        graph_window = window

if not graph_window:
    raise AttributeError("Could not find Graph tree reordering window.")

# initialise the btn variables
btn_move_up = None
# btn_move_down = None
btn_ok = None
btn_apply = None
list_box = None
# loop through Graph tree reordering window and find the buttons we need and
# the list_box.

for child in graph_window.children():
    print(child.window_text())


for child in graph_window.children():
    if child.window_text() == window_text_ok:
        btn_ok = child
    if child.window_text() == window_text_apply:
        btn_apply = child
    if child.window_text() == window_text_move_up:
        btn_move_up = child
    if child.window_text() == window_text_move_down:
        btn_move_down = child
    if child.window_text() == window_text_list_box:
        list_box = child

# create a text list of the items in the list box and sort them.
tree_items = []
for text in list_box.item_texts():  # type: ignore
    tree_items.append(text)
tree_items.sort()

# reorder the items in the tree
for i, value in enumerate(tree_items):
    # select the item in list box
    list_box.select(value)  # type: ignore
    # find it's current position in the list_box
    current_position = list_box.item_texts().index(value)  # type: ignore
    target_position = i
    # current_position - target position is the number of times to click the
    # move up button.
    for _ in range(current_position - target_position):
        btn_move_up.click()  # type: ignore

btn_apply.click()  # type: ignore
btn_ok.click()  # type: ignore
