#! /usr/bin/python3.6

"""

    Example 15:

    Draw a line between two points.

"""
from pycatia import catia

documents = catia.documents
# create a new CATPart.
documents.add('Part')
document = catia.active_document

part = document.part()
hybrid_bodies = part.hybrid_bodies
hsf = part.hybrid_shape_factory

# create a new hybrid body.
geom_set = hybrid_bodies.add()
geom_set.name = 'Construction_Geometry'

co_ord_1 = (0, 0, 0)
co_ord_2 = (100, 0, 0)
point_1 = hsf.add_new_point_coord(co_ord_1[0], co_ord_1[1], co_ord_1[2])
point_2 = hsf.add_new_point_coord(co_ord_2[0], co_ord_2[1], co_ord_2[2])

geom_set.append_hybrid_shape(point_1)
geom_set.append_hybrid_shape(point_2)

line = hsf.add_new_line_pt_pt(point_1, point_2)

geom_set.append_hybrid_shape(line)

part.update()
