"""Functions for metrics related to 2D and 3D bounding boxes."""

# pylint: disable=invalid-name,missing-docstring,assignment-from-no-return,logging-fstring-interpolation

import logging

import numpy as np

from utils.bbox.geometry import polygon_area, polygon_collision, polygon_intersection

from .bbox2d import BBox2D
from .bbox2d_list import BBox2DList

logger = logging.getLogger(__name__)


def iou_2d(a: BBox2D, b: BBox2D):
    """
    Compute the Intersection over Union (IoU) of a pair of 2D bounding boxes.

    Alias for `jaccard_index_2d`.
    """
    return jaccard_index_2d(a, b)


def is_cross(a: BBox2D, b: BBox2D):
    ax1, ay1, ax2, ay2 = a.x1, a.y1, a.x2, a.y2  # прямоугольник А
    bx1, by1, bx2, by2 = b.x1, b.y1, b.x2, b.y2  # прямоугольник B
    # это были координаты точек диагонали по каждому прямоугольнику

    # 1. Проверить условия перекрытия, например, если XПA<XЛB ,
    #    то прямоугольники не пересекаются,и общая площадь равна нулю.
    #   (это случай, когда они справа и слева) и аналогично, если они сверху
    #    и снизу относительно друг друга.
    #    (XПА - это  Х Правой точки прямоугольника А)
    #    (ХЛВ - Х Левой точки прямоугольника В )
    #    нарисуй картинку (должно стать понятнее)

    xA = [ax1, ax2]  # координаты x обеих точек прямоугольника А
    xB = [bx1, bx2]  # координаты x обеих точке прямоугольника В

    yA = [ay1, ay2]  # координаты x обеих точек прямоугольника А
    yB = [by1, by2]  # координаты x обеих точек прямоугольника В

    if max(xA) < min(xB) or max(yA) < min(yB) or min(yA) > max(yB) or min(xA) > max(xB):
        return False  # не пересекаются

    return True


def jaccard_index_2d(a: BBox2D, b: BBox2D):
    """
    Compute the Jaccard Index / Intersection over Union (IoU) of a pair of 2D bounding boxes.

    Args:
        a (:py:class:`BBox2D`): 2D bounding box.
        b (:py:class:`BBox2D`): 2D bounding box.

    Returns:
        :py:class:`float`: The IoU of the 2 bounding boxes.
    """

    xA = np.maximum(a.x1, b.x1)
    yA = np.maximum(a.y1, b.y1)
    xB = np.minimum(a.x2, b.x2)
    yB = np.minimum(a.y2, b.y2)

    logger.debug("xA={xA} yA={yA} xB={xB} yB={yB}")

    inter_w = xB - xA + 1
    inter_w = inter_w * (inter_w >= 0)

    inter_h = yB - yA + 1
    inter_h = inter_h * (inter_h >= 0)

    intersection = inter_w * inter_h

    logger.debug(f"jaccard_index: intersection={intersection}")

    a_area = a.width * a.height
    b_area = b.width * b.height

    logger.debug(f"jaccard_index: a_area: {a_area}, b_area: {b_area}")

    iou = intersection / (a_area + b_area - intersection)

    # set nan and +/- inf to 0
    if np.isinf(iou) or np.isnan(iou):
        iou = 0

    return iou


def multi_iou_2d(a: BBox2DList, b: BBox2DList):
    """
    Compute the Intersection over Union (IoU) of two sets of 2D bounding boxes.

    Alias for `multi_jaccard_index_2d`.
    """
    return multi_jaccard_index_2d(a, b)


def multi_jaccard_index_2d(a: BBox2DList, b: BBox2DList):
    """
    Compute the Jaccard Index (Intersection over Union) of two sets of 2D bounding boxes.

    Args:
        a (:py:class:`BBox2DList`): List of 2D bounding boxes.
        b (:py:class:`BBox2DList`): List of 2D bounding boxes.

    Returns:
        :py:class:`ndarray`: IoU Matrix
    """

    # We need to add a trailing dimension so that max/min gives us a (N,N) matrix
    xA = np.maximum(a.x1[:, np.newaxis], b.x1[:, np.newaxis].T)
    yA = np.maximum(a.y1[:, np.newaxis], b.y1[:, np.newaxis].T)
    xB = np.minimum(a.x2[:, np.newaxis], b.x2[:, np.newaxis].T)
    yB = np.minimum(a.y2[:, np.newaxis], b.y2[:, np.newaxis].T)

    logger.debug(
        "\nmulti_jaccard_index:\nxA\n{xA}\nyA\n{yA}\nxB\n{xB}\nyB\n{yB}")

    inter_w = xB - xA + 1
    inter_w[inter_w < 0] = 0

    inter_h = yB - yA + 1
    inter_h[inter_h < 0] = 0

    # maximum generates a (N,N) matrix which consumes a lot of memory
    # thus we are aggressive about freeing memory up.
    del xA
    del yA
    del xB
    del yB

    intersection = inter_w * inter_h
    logger.debug(f"\nmulti_jaccard_index intersection:\n {intersection}")

    a_area = a.width[:, np.newaxis] * a.height[:, np.newaxis]
    b_area = b.width[:, np.newaxis] * b.height[:, np.newaxis]
    logger.debug(
        f"\nmulti_jaccard_index:\n a_area:\n {a_area} \nb_area:\n {b_area}")

    iou = intersection / (a_area + b_area.T - intersection)

    # set nan and +/- inf to 0
    iou[np.isinf(iou)] = 0
    iou[np.isnan(iou)] = 0

    return iou
