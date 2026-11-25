"""
VDOT LRS Event -> MicroStation alignment
========================================

Connects to VDOT's public LRS Route Master REST service, lets you choose a
route from drop-down menus, previews the route and the event in the
MicroStation view, and lets you pick the From and To measures on the route
with data points. "Create alignment" then draws the event along the real
route centreline and attaches the LRS metadata as an Item Type.

Service:  LRS_Route_Master (ArcGIS FeatureServer, polylines with M = miles)

Run it from the Python Manager (Utilities > Python) in MicroStation 2026 or
OpenRoads Designer 2025, or key in:  PYTHON LOAD "<folder>/VDOT_LRS_Linear_Milepost.py"
The VDOT logo is embedded in this file, so the script is self-contained. In
OpenRoads the result is plain DGN
graphics (line strings, cells, text), not civil geometry.

In the drawing
--------------
* The whole route is drawn thin, the event thick, From as a green dot and To
  as a red dot. The preview is tool dynamics: it shows while the pick tool is
  active and the cursor is over a view.
* Data point  = set the measure nearest the cursor (From first, then To).
* Reset       = leave the pick tool ("Pick From / To in drawing" restarts it).
* Preview colours are colour-table indexes (see PREVIEW_* below).

Notes
-----
* Latitude/longitude from the service are WGS84. They are converted with the
  model's Geographic Coordinate System, an EPSG system chosen in the dialog,
  or a local tangent plane when no GCS is wanted. WGS84 is used directly in
  the GCS datum (NAD83); the difference is about 1 m.
* A 2D model cannot hold 3D elements: elements are 3D only in a 3D model.
* An event that crosses a gap between route parts becomes one element per part.
* "Milepost signs" adds green reference-location signs (MILE, stacked whole-mile
  digits, a rule and the decimal) at From/To or at a regular spacing. Each sign
  is one cell named MILEPOST; its size follows the spacing (see MILEPOST_SIZE).
* Labels sit to the left of the direction of travel and milepost signs to the
  right, moved out until they are clear of the centreline. Their size follows the event length
  but is kept within LABEL_HEIGHT_M and SIGN_HEIGHT_M (metres).
* Service requests run on a background thread so MicroStation stays responsive;
  everything that touches the drawing stays on the main thread. Route lists and
  routes already downloaded are reused while the dialog stays open.
* The theme, jurisdiction, coordinate system, label and milepost choices and the
  log being open or closed are remembered between runs.
"""

import base64
import bisect
import html
import json
import math
import queue
import sys
import tempfile
import threading
import time
import traceback
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass

try:
    from MSPyBentley import *
    from MSPyBentleyGeom import *
    from MSPyECObjects import *
    from MSPyDgnPlatform import *
    from MSPyDgnView import *
    from MSPyMstnPlatform import *
    IN_MSTN = True
except ImportError:
    IN_MSTN = False

def _use_bundled_libs():
    """Fall back to the PyQt5 the installer put in lib\\ beside this script.

    The PyQt5 that ships with MicroStation/OpenRoads is always preferred; the bundled copy
    is appended to the end of sys.path, so it is used only when that one is missing.
    """
    import os
    try:
        lib = os.path.join(os.path.dirname(os.path.abspath(__file__)), "lib")
    except NameError:
        return False
    if os.path.isdir(lib) and lib not in sys.path:
        sys.path.append(lib)
        return True
    return False


for _attempt in range(2):
    try:
        from PyQt5.QtCore import Qt, QTimer, QPointF, QSettings
        from PyQt5.QtGui import QColor, QPainter, QPixmap, QPolygonF
        from PyQt5.QtWidgets import (
            QApplication, QWidget, QFrame, QLabel, QComboBox, QPushButton, QButtonGroup,
            QTextEdit, QGridLayout, QVBoxLayout, QHBoxLayout, QScrollArea, QCompleter,
            QMessageBox,
        )
        HAVE_QT = True
        break
    except ImportError:
        HAVE_QT = False
        if _attempt or not _use_bundled_libs():
            break


SERVICE_URL = ("https://services.arcgis.com/p5v98VHDX9Atv3l7/arcgis/rest/services/"
               "LRS_Route_Master/FeatureServer/0")

# Selections made when the dialog opens. None leaves a drop-down blank. To preset one,
# use the exact service values, e.g. "Washington County" and "R-VA095SC00700EB".
DEFAULT_JURISDICTION = None
DEFAULT_ROUTE_ID = None
STATEWIDE_LABEL = "Statewide routes (Interstate, US, VA)"
DEFAULT_FROM_MEASURE = 0.000
DEFAULT_TO_MEASURE = 15.810

APP_TITLE = "VDOT LRS Linear Milepost"
APP_AUTHOR = "Robert Ratliff"
BUILD_DATE = "2026-11-25"

ITEM_LIBRARY = "VDOT LRS"
ITEM_TYPE = "Linear Event"
# Labels, start/stop markers and milepost signs carry this item type, so "Delete" can
# find everything the tool placed, in this session or a later one.
GRAPHIC_TYPE = "LRS Graphic"
GRAPHIC_PART = "Part"
PART_LABEL, PART_FROM, PART_TO, PART_MILEPOST = "Label", "Start marker", "Stop marker", "Milepost"

EPSG_CHOICES = (
    (2284, "NAD83 Virginia South (ftUS)"),
    (2283, "NAD83 Virginia North (ftUS)"),
    (6595, "NAD83(2011) Virginia South (ftUS)"),
    (6593, "NAD83(2011) Virginia North (ftUS)"),
    (3968, "NAD83 Virginia Lambert (m)"),
)

MODE_MODEL_GCS = "model"
MODE_EPSG = "epsg"
MODE_LOCAL = "local"

EVENT_LINEAR = "Linear"
EVENT_POINT = "Point"

LABELS_ALL = "all"
LABELS_MEASURES = "measures"
LABELS_NONE = "none"

MILEPOST_NONE = "none"
MILEPOST_ENDS = "ends"
MILEPOST_CHOICES = (
    ("No milepost signs", MILEPOST_NONE),
    ("At From and To", MILEPOST_ENDS),
    ("Every mile", 1.0),
    ("Every 0.5 mile", 0.5),
    ("Every 0.1 mile", 0.1),
)
MAX_MILEPOSTS = 400
MILEPOST_GREEN = (0, 110, 81)        # sign background (RGB)
MILEPOST_WHITE = (255, 255, 255)     # border and legend (RGB)
MILEPOST_SIZE = 1.0                  # multiplier on the automatic sign size
LABEL_HEIGHT_M = (5.0, 50.0)         # smallest and largest label text height, metres
SIGN_HEIGHT_M = (20.0, 250.0)        # smallest and largest milepost sign height, metres
LABEL_CHAR_WIDTH = 0.7               # estimated character width as a share of text height

# Preview symbology: (colour-table index, weight). Default table: 1 blue, 2 green,
# 3 red, 4 yellow, 7 cyan.
PREVIEW_ROUTE = (1, 1)
PREVIEW_EVENT = (7, 5)
PREVIEW_FROM = (2, 30)       # weight 30 is close to MicroStation's maximum of 31
PREVIEW_TO = (3, 30)
PREVIEW_CURSOR = (4, 16)

# Symbology of what "Create alignment" adds besides the line itself.
CREATED_FROM = (2, 15)       # start marker - green
CREATED_TO = (3, 15)         # stop marker - red
POINT_EVENT = (4, 15)        # Point event type - yellow (preview and created element)

MAX_LINESTRING_POINTS = 4500
WGS84_A = 6378137.0
WGS84_E2 = 0.00669437999014


def fmt_measure(measure):
    return f"{round(measure, 3) + 0.0:.3f}"


# ---------------------------------------------------------------------------
# Route geometry with measures (pure Python)
# ---------------------------------------------------------------------------
class RouteGeometry:
    """Polyline parts of (x, y, m) vertices with m increasing along each part."""

    def __init__(self, paths):
        self.paths = []
        for path in paths:
            vertices = [(v[0], v[1], v[2]) for v in path if len(v) > 2 and v[2] is not None]
            if len(vertices) < 2:
                continue
            if vertices[0][2] > vertices[-1][2]:
                vertices.reverse()
            self.paths.append(vertices)
        if not self.paths:
            raise ValueError("The route has no geometry with measures.")
        self.paths.sort(key=lambda p: p[0][2])
        self._measures = [[v[2] for v in path] for path in self.paths]
        self.m_min = self.paths[0][0][2]
        self.m_max = max(path[-1][2] for path in self.paths)
        self._measure_per_unit = None

    @property
    def measure_per_unit(self):
        """Average measure covered by one unit of length along the route."""
        if self._measure_per_unit is None:
            length = sum(polyline_length(path) for path in self.paths)
            covered = sum(measures[-1] - measures[0] for measures in self._measures)
            self._measure_per_unit = covered / length if length > 0.0 else 0.0
        return self._measure_per_unit

    def window(self, m_from, m_to):
        """Vertices of each part around a range of measures, as lists of (x, y, m)."""
        parts = []
        for path, measures in zip(self.paths, self._measures):
            first = max(bisect.bisect_left(measures, m_from) - 1, 0)
            last = min(bisect.bisect_right(measures, m_to) + 1, len(path))
            if last - first >= 2:
                parts.append(path[first:last])
        return parts

    def mapped(self, transform):
        return RouteGeometry(
            [[transform(x, y) + (m,) for x, y, m in path] for path in self.paths])

    @staticmethod
    def _interpolate(path, measures, m):
        index = min(max(bisect.bisect_right(measures, m), 1), len(path) - 1)
        ax, ay, am = path[index - 1]
        bx, by, bm = path[index]
        t = 0.0 if bm == am else min(max((m - am) / (bm - am), 0.0), 1.0)
        return ax + (bx - ax) * t, ay + (by - ay) * t

    def point_at(self, m):
        """Location of a measure; a measure in a gap snaps to the nearest part end."""
        for path, measures in zip(self.paths, self._measures):
            if measures[0] <= m <= measures[-1]:
                return self._interpolate(path, measures, m)
        ends = [path[i] for path in self.paths for i in (0, -1)]
        x, y, _m = min(ends, key=lambda v: abs(v[2] - m))
        return x, y

    def direction_at(self, m):
        """Unit vector along the route (towards higher measures) at a measure."""
        chosen = None
        for path, measures in zip(self.paths, self._measures):
            if measures[0] <= m <= measures[-1]:
                chosen = (path, measures)
                break
        if chosen is None:
            chosen = min(zip(self.paths, self._measures),
                         key=lambda pm: min(abs(pm[1][0] - m), abs(pm[1][-1] - m)))
        path, measures = chosen
        index = min(max(bisect.bisect_right(measures, m), 1), len(path) - 1)
        for offset in range(len(path)):            # skips zero-length segments
            for i in (index + offset, index - offset):
                if 1 <= i < len(path):
                    dx, dy = path[i][0] - path[i - 1][0], path[i][1] - path[i - 1][1]
                    length = math.hypot(dx, dy)
                    if length > 0.0:
                        return dx / length, dy / length
        return 1.0, 0.0

    def slice(self, m_from, m_to):
        """Parts of the route between two measures, as lists of (x, y)."""
        parts = []
        for path, measures in zip(self.paths, self._measures):
            low, high = max(m_from, measures[0]), min(m_to, measures[-1])
            if high <= low:
                continue
            points = [self._interpolate(path, measures, low)]
            points += [(x, y) for x, y, m in path if low < m < high]
            points.append(self._interpolate(path, measures, high))
            parts.append(points)
        return parts

    def locate(self, x, y):
        """Nearest location on the route: (measure, x, y)."""
        best = None
        for path in self.paths:
            ax, ay, am = path[0]
            for bx, by, bm in path[1:]:
                dx, dy = bx - ax, by - ay
                length2 = dx * dx + dy * dy
                t = 0.0 if length2 == 0.0 else min(max(((x - ax) * dx + (y - ay) * dy) / length2, 0.0), 1.0)
                px, py = ax + dx * t, ay + dy * t
                distance2 = (x - px) ** 2 + (y - py) ** 2
                if best is None or distance2 < best[0]:
                    best = (distance2, am + (bm - am) * t, px, py)
                ax, ay, am = bx, by, bm
        return best[1], best[2], best[3]

    def extent(self):
        xs = [v[0] for path in self.paths for v in path]
        ys = [v[1] for path in self.paths for v in path]
        return min(xs), min(ys), max(xs), max(ys)

    def has_measure(self, m):
        return any(ms[0] - 1e-6 <= m <= ms[-1] + 1e-6 for ms in self._measures)

    def stops(self):
        """Measures offered in the drop-downs: part ends and every whole mile."""
        values = set()
        for measures in self._measures:
            values.add(round(measures[0], 3) + 0.0)
            values.add(round(measures[-1], 3) + 0.0)
            for mile in range(math.ceil(measures[0]), math.floor(measures[-1]) + 1):
                values.add(float(mile))
        return sorted(values)


def polyline_length(points):
    return sum(math.hypot(b[0] - a[0], b[1] - a[1]) for a, b in zip(points, points[1:]))


def clamp(value, low, high):
    return min(max(value, low), high)


def side_centre(x, y, direction, side, width, height, gap):
    """Centre of an upright width x height box placed beside the route at (x, y).

    side is +1 for the right of the direction of travel, -1 for the left; gap is the
    clear distance between the route and the nearest part of the box.
    """
    nx, ny = direction[1] * side, -direction[0] * side
    reach = gap + abs(nx) * width / 2.0 + abs(ny) * height / 2.0
    return x + nx * reach, y + ny * reach


def segment_hits_box(ax, ay, bx, by, x_min, y_min, x_max, y_max):
    """True when the segment a-b touches the axis-aligned box (Liang-Barsky clipping)."""
    if (max(ax, bx) < x_min or min(ax, bx) > x_max
            or max(ay, by) < y_min or min(ay, by) > y_max):
        return False
    dx, dy = bx - ax, by - ay
    enter, leave = 0.0, 1.0
    for p, q in ((-dx, ax - x_min), (dx, x_max - ax), (-dy, ay - y_min), (dy, y_max - ay)):
        if p == 0.0:
            if q < 0.0:
                return False
        elif p < 0.0:
            if q / p > leave:
                return False
            enter = max(enter, q / p)
        else:
            if q / p < enter:
                return False
            leave = min(leave, q / p)
    return True


def beside_route(geometry, measure, side, width, height, gap, stack=0.0):
    """Centre of an upright width x height box beside the route at a measure.

    The box starts `gap` clear of the route (see side_centre) and, because a road that
    curves can swing back into it, is moved further out in steps until it no longer
    touches the route nearby. If no position on that side is clear (a sharp bend), the
    other side is tried; failing that the box stays at its first position. `stack`
    shifts it vertically (a second line of text).
    """
    x, y = geometry.point_at(measure)
    direction = geometry.direction_at(measure)
    step = 0.75 * min(width, height)
    span = 1.5 * (width + height + gap + 12 * step + abs(stack)) * geometry.measure_per_unit
    nearby = geometry.window(measure - span, measure + span)
    margin = gap / 2.0                       # the box must clear the route by this much
    first = None
    for try_side in (side, -side):
        for extra in range(13):
            centre_x, centre_y = side_centre(x, y, direction, try_side, width, height,
                                             gap + extra * step)
            centre_y += stack
            if first is None:
                first = (centre_x, centre_y)
            box = (centre_x - width / 2.0 - margin, centre_y - height / 2.0 - margin,
                   centre_x + width / 2.0 + margin, centre_y + height / 2.0 + margin)
            if not any(segment_hits_box(a[0], a[1], b[0], b[1], *box)
                       for part in nearby for a, b in zip(part, part[1:])):
                return centre_x, centre_y
    return first


def milepost_text(measure, mode):
    """Sign legend for a measure: (whole-mile digits, decimal such as '.5' or '')."""
    value = max(round(measure, 3), 0.0) + 0.0      # "+ 0.0" turns -0.0 into 0.0
    whole = int(value)
    decimal = f"{value - whole:.3f}"[1:].rstrip("0")
    if decimal == ".":
        decimal = "" if mode in (1.0, MILEPOST_ENDS) else ".0"
    return str(whole), decimal


def milepost_layout(digits, decimal, unit):
    """Sign face for a legend, with y measured up from the bottom edge.

    Returns (width, height, [(text, centre y, text height)], rule y or None).
    `unit` is the height of one digit.
    """
    margin, gap, mile_height = 0.45 * unit, 0.35 * unit, 0.5 * unit
    rows = [("MILE", mile_height)] + [(digit, unit) for digit in digits]
    height = 2 * margin + sum(h for _t, h in rows) + gap * (len(rows) - 1)
    if decimal:
        height += 2 * gap + unit
    texts, y = [], height - margin
    for text, row_height in rows:
        texts.append((text, y - row_height / 2.0, row_height))
        y -= row_height + gap
    rule = None
    if decimal:
        rule = y
        texts.append((decimal, y - gap - unit / 2.0, unit))
    return 3.0 * unit, height, texts, rule


def rounded_rectangle(x_centre, bottom, width, height, radius, steps=6):
    """Closed outline of a rounded rectangle as (x, y) points."""
    left, right, top = x_centre - width / 2.0, x_centre + width / 2.0, bottom + height
    corners = ((right - radius, bottom + radius, -90.0), (right - radius, top - radius, 0.0),
               (left + radius, top - radius, 90.0), (left + radius, bottom + radius, 180.0))
    points = []
    for cx, cy, start in corners:
        for step in range(steps + 1):
            angle = math.radians(start + 90.0 * step / steps)
            points.append((cx + radius * math.cos(angle), cy + radius * math.sin(angle)))
    points.append(points[0])
    return points


def local_xy_m(lon, lat, lon0, lat0):
    """Metres east/north of (lon0, lat0) on a tangent plane at that point."""
    phi = math.radians(lat0)
    w = math.sqrt(1.0 - WGS84_E2 * math.sin(phi) ** 2)
    east = math.radians(lon - lon0) * (WGS84_A / w) * math.cos(phi)
    north = math.radians(lat - lat0) * (WGS84_A * (1.0 - WGS84_E2) / w ** 3)
    return east, north


# ---------------------------------------------------------------------------
# VDOT LRS REST service
# ---------------------------------------------------------------------------
class LrsServiceError(Exception):
    pass


class LrsOffline(LrsServiceError):
    """The service could not be reached at all."""


@dataclass
class RouteInfo:
    name_id: str
    common_name: str
    jurisdiction: str
    category: str
    direction: str
    opposite_name: str
    geometry: RouteGeometry          # x = longitude, y = latitude (WGS84)
    # Most undivided roads name an opposite direction that is not a route of its own.
    opposite_exists: bool = False

    @property
    def route(self):
        return self.common_name.split(" (")[0]


class LrsService:
    def __init__(self, url=SERVICE_URL, timeout=45):
        self.url = url
        self.timeout = timeout
        self.on_retry = None                   # optional callback(attempt number)
        # Replies already downloaded by this object; a failed request is never stored.
        self._jurisdictions = None
        self._routes_cache = {}
        self._route_cache = {}

    def cached_routes(self, jurisdiction):
        return self._routes_cache.get(jurisdiction)

    def cached_route(self, name_id):
        return self._route_cache.get(name_id)

    def _query(self, **params):
        params.setdefault("f", "json")
        request = urllib.request.Request(
            self.url + "/query", data=urllib.parse.urlencode(params).encode("ascii"),
            headers={"User-Agent": "MicroStation-VDOT-LRS/1.0"})
        payload = None
        for attempt in range(3):               # a dropped connection is retried twice
            try:
                with urllib.request.urlopen(request, timeout=self.timeout) as response:
                    payload = json.loads(response.read().decode("utf-8"))
                break
            except (urllib.error.URLError, TimeoutError, OSError) as ex:
                if isinstance(ex, urllib.error.HTTPError) or attempt == 2:
                    raise LrsOffline(
                        f"Could not reach the VDOT LRS service: {ex}") from ex
                if self.on_retry is not None:
                    self.on_retry(attempt + 1)
                time.sleep(1.0 + attempt)
            except ValueError as ex:
                raise LrsServiceError(
                    "The VDOT LRS service returned an unreadable reply.") from ex
        if "error" in payload:
            error = payload["error"]
            raise LrsServiceError(
                f"VDOT LRS service error {error.get('code', '')}: {error.get('message', 'unknown')}")
        return payload

    def _features(self, **params):
        """All features of a query, following the service's paging."""
        features, offset = [], 0
        while True:
            payload = self._query(resultOffset=offset, **params)
            page = payload.get("features", [])
            features += page
            if not payload.get("exceededTransferLimit") or not page:
                return features
            offset += len(page)

    def _attributes(self, **params):
        """All rows of a table query, following the service's paging."""
        return [feature["attributes"]
                for feature in self._features(returnGeometry="false", **params)]

    @staticmethod
    def _core_extent(points):
        """Bounding box (xmin, ymin, xmax, ymax) of points, ignoring far-off strays.

        A few records carry the wrong jurisdiction (one "City of Richmond" street lies
        near Roanoke); a plain bounding box would then span half the state.
        """
        if not points:
            return None

        def span(values):
            values = sorted(values)
            low = values[int(0.02 * (len(values) - 1))]
            high = values[int(round(0.98 * (len(values) - 1)))]
            pad = max(0.5 * (high - low), 0.02)            # 0.02 degrees is about a mile
            kept = [v for v in values if low - pad <= v <= high + pad]
            return kept[0], kept[-1]

        xmin, xmax = span([p[0] for p in points])
        ymin, ymax = span([p[1] for p in points])
        return xmin, ymin, xmax, ymax

    @staticmethod
    def _quote(text):
        return "'" + text.replace("'", "''") + "'"

    def _jurisdiction_where(self, jurisdiction):
        if jurisdiction:
            return f"RTE_JURIS_PROPER_NM = {self._quote(jurisdiction)}"
        return self.STATEWIDE

    def jurisdictions(self):
        if self._jurisdictions is None:
            rows = self._attributes(where="1=1", outFields="RTE_JURIS_PROPER_NM",
                                    returnDistinctValues="true",
                                    orderByFields="RTE_JURIS_PROPER_NM")
            self._jurisdictions = sorted(
                {(row["RTE_JURIS_PROPER_NM"] or "").strip() for row in rows})
        return self._jurisdictions

    # Interstates, US routes and most VA primaries are single statewide records with no
    # jurisdiction, so they are added to a jurisdiction's list by location instead.
    STATEWIDE = "(RTE_JURIS_PROPER_NM IS NULL OR RTE_JURIS_PROPER_NM = '')"
    NOT_ROUTES = "RTE_CATEGORY_NM NOT LIKE 'RNS Datum%'"      # internal LRS layers
    HIGHWAY_ORDER = {"Interstate": 0, "US Highway Primary": 1, "State Highway Primary": 2}
    ROUTE_FIELDS = "RTE_NM,RTE_COMMON_NM,RTE_NBR,RTE_CATEGORY_NM"

    @staticmethod
    def _route_items(rows):
        return [(row["RTE_COMMON_NM"] or row["RTE_NM"], row["RTE_NM"]) for row in rows]

    def _highways_first(self, rows):
        return sorted(rows, key=lambda row: (
            self.HIGHWAY_ORDER.get(row["RTE_CATEGORY_NM"], 3),
            row["RTE_NBR"] if row["RTE_NBR"] is not None else 10 ** 9,
            row["RTE_COMMON_NM"] or row["RTE_NM"]))

    def routes(self, jurisdiction):
        """Routes offered for a jurisdiction: ([(common name, route name ID)], statewide count).

        The statewide routes (Interstate, US, VA, then ramps and frontage roads) that
        cross the jurisdiction's bounding box come first, then the jurisdiction's own
        routes in route-number order. An empty jurisdiction lists every statewide route.
        """
        if jurisdiction not in self._routes_cache:
            self._routes_cache[jurisdiction] = self._download_routes(jurisdiction)
        return self._routes_cache[jurisdiction]

    def _download_routes(self, jurisdiction):
        if not jurisdiction:
            rows = self._highways_first(self._attributes(
                where=f"{self.STATEWIDE} AND {self.NOT_ROUTES}", outFields=self.ROUTE_FIELDS,
                orderByFields="RTE_NBR,RTE_COMMON_NM"))
            return self._route_items(rows), len(rows)

        # The jurisdiction's routes come back with heavily simplified geometry, used only
        # to work out the area the jurisdiction covers.
        features = self._features(
            where=self._jurisdiction_where(jurisdiction), outFields=self.ROUTE_FIELDS,
            orderByFields="RTE_NBR,RTE_COMMON_NM", returnGeometry="true", outSR="4326",
            maxAllowableOffset="0.05", geometryPrecision="4")
        local = [feature["attributes"] for feature in features]
        points = [path[index][:2] for feature in features
                  for path in (feature.get("geometry") or {}).get("paths", [])
                  for index in (0, -1) if path]
        statewide = []
        extent = self._core_extent(points)
        if extent is not None:
            statewide = self._highways_first(self._attributes(
                where=f"{self.STATEWIDE} AND {self.NOT_ROUTES}", outFields=self.ROUTE_FIELDS,
                geometry=",".join(str(value) for value in extent),
                geometryType="esriGeometryEnvelope", inSR="4326",
                spatialRel="esriSpatialRelIntersects",
                orderByFields="RTE_NBR,RTE_COMMON_NM"))
        return self._route_items(statewide + local), len(statewide)

    def route(self, name_id):
        if name_id not in self._route_cache:
            self._route_cache[name_id] = self._download_route(name_id)
        return self._route_cache[name_id]

    def _download_route(self, name_id):
        payload = self._query(
            where=f"RTE_NM = {self._quote(name_id)}", returnGeometry="true", returnM="true",
            outSR="4326",
            outFields="RTE_NM,RTE_COMMON_NM,RTE_JURIS_PROPER_NM,RTE_CATEGORY_NM,"
                      "RTE_DIRECTION_CD,RTE_OPPOSITE_DIRECTION_RTE_NM")
        features = payload.get("features", [])
        if not features:
            raise LrsServiceError(f"Route {name_id} was not found in the LRS service.")
        paths = []
        for feature in features:
            paths += (feature.get("geometry") or {}).get("paths", [])
        attributes = features[0]["attributes"]
        try:
            geometry = RouteGeometry(paths)
        except ValueError as ex:
            raise LrsServiceError(f"Route {name_id}: {ex}") from ex
        opposite = attributes.get("RTE_OPPOSITE_DIRECTION_RTE_NM") or ""
        opposite_exists = bool(opposite) and self._query(
            where=f"RTE_NM = {self._quote(opposite)}", returnCountOnly="true").get("count", 0) > 0
        return RouteInfo(
            name_id=attributes["RTE_NM"],
            common_name=attributes.get("RTE_COMMON_NM") or attributes["RTE_NM"],
            jurisdiction=attributes.get("RTE_JURIS_PROPER_NM") or "",
            category=attributes.get("RTE_CATEGORY_NM") or "",
            direction=attributes.get("RTE_DIRECTION_CD") or "",
            opposite_name=opposite, geometry=geometry, opposite_exists=opposite_exists)


# ---------------------------------------------------------------------------
# Event
# ---------------------------------------------------------------------------
@dataclass
class LrsEvent:
    route_info: RouteInfo
    event_type: str
    from_measure: float
    to_measure: float

    def validate(self):
        geometry = self.route_info.geometry
        low, high = geometry.m_min - 0.0005, geometry.m_max + 0.0005
        measures = [("From", self.from_measure)]
        if self.event_type == EVENT_LINEAR:
            measures.append(("To", self.to_measure))
        for label, measure in measures:
            if not low <= measure <= high:
                raise ValueError(
                    f"{label} measure {fmt_measure(measure)} is outside the route "
                    f"({fmt_measure(geometry.m_min)} to {fmt_measure(geometry.m_max)} mi).")
        if self.event_type == EVENT_LINEAR and self.to_measure <= self.from_measure:
            raise ValueError("To measure must be greater than From measure.")

    @property
    def length(self):
        return self.to_measure - self.from_measure if self.event_type == EVENT_LINEAR else 0.0

    def lat_lon(self, measure):
        lon, lat = self.route_info.geometry.point_at(measure)
        return lat, lon

    def milepost_measures(self, mode):
        """Measures that get a milepost sign for a MILEPOST_CHOICES mode."""
        if mode == MILEPOST_NONE:
            return []
        if self.event_type == EVENT_POINT:
            return [self.from_measure]
        if mode == MILEPOST_ENDS:
            return [self.from_measure, self.to_measure]
        step = round(mode * 10)
        first = math.ceil(self.from_measure * 10 - 1e-6)
        last = math.floor(self.to_measure * 10 + 1e-6)
        geometry = self.route_info.geometry
        return [tenth / 10.0 for tenth in range(first, last + 1)
                if tenth % step == 0 and geometry.has_measure(tenth / 10.0)]


# ---------------------------------------------------------------------------
# MicroStation side
# ---------------------------------------------------------------------------
def active_model():
    model_ref = ISessionMgr.ActiveDgnModelRef
    if model_ref is None:
        raise RuntimeError("There is no active DGN model.")
    return model_ref, model_ref.GetDgnModel()


def model_summary():
    """(is3d, GCS name or None) of the active model."""
    model_ref, dgn_model = active_model()
    gcs = DgnGCS.FromModel(model_ref, True)
    return dgn_model.Is3d(), (str(gcs.GetName()) or "unnamed GCS") if gcs is not None else None


def make_projector(mode, epsg, origin, log):
    """Return f(lon, lat) -> (x, y) in UORs of the active model."""
    model_ref, dgn_model = active_model()

    if mode == MODE_LOCAL:
        uor_per_meter = dgn_model.GetModelInfo().GetUorPerMeter()
        lon0, lat0 = origin
        log("Local tangent plane: the route start is placed at the model origin.")

        def project_local(lon, lat):
            east, north = local_xy_m(lon, lat, lon0, lat0)
            return east * uor_per_meter, north * uor_per_meter
        return project_local

    if mode == MODE_MODEL_GCS:
        gcs = DgnGCS.FromModel(model_ref, True)
        if gcs is None:
            raise RuntimeError("The active model has no Geographic Coordinate System.")
        log(f"Converting with the model GCS '{gcs.GetName()}'.")
    else:
        base = BaseGCS()
        if base.InitFromEPSGCode(int(epsg))[0] != 0:
            raise RuntimeError(f"EPSG:{epsg} was not found in the GCS library.")
        gcs = DgnGCS.CreateGCS(base, model_ref)
        if gcs is None:
            raise RuntimeError(f"Could not create a GCS from EPSG:{epsg}.")
        log(f"Converting with EPSG:{epsg} ({gcs.GetName()}).")

    accepted = (ReprojectStatus.eREPROJECT_Success,
                ReprojectStatus.eREPROJECT_CSMAPERR_OutOfUsefulRange)
    geo, uors = GeoPoint(), DPoint3d()

    def project_gcs(lon, lat):
        geo.Init(lon, lat, 0.0)
        status = gcs.UorsFromLatLong(uors, geo)
        if status not in accepted:
            raise RuntimeError(f"({lat:.6f}, {lon:.6f}) could not be converted: {status}.")
        return uors.x, uors.y
    return project_gcs


def _apply_symbology(eeh, symbology):
    setter = ElementPropertiesSetter()
    setter.SetColor(symbology[0])
    setter.SetWeight(symbology[1])
    setter.Apply(eeh)


def linestring_elements(points, symbology=None):
    """Line string element(s) through (x, y) points, split at the vertex limit."""
    model_ref, dgn_model = active_model()
    elements = []
    start = 0
    while start < len(points) - 1:
        chunk = points[start:start + MAX_LINESTRING_POINTS]
        eeh = EditElementHandle()
        array = DPoint3dArray([DPoint3d(x, y, 0.0) for x, y in chunk])
        if BentleyStatus.eSUCCESS != LineStringHandler.CreateLineStringElement(
                eeh, None, array, dgn_model.Is3d(), model_ref):
            raise RuntimeError("LineStringHandler could not create an element.")
        if symbology is not None:
            _apply_symbology(eeh, symbology)
        elements.append(eeh)
        start += MAX_LINESTRING_POINTS - 1
    return elements


def point_element(x, y, symbology=None):
    """Zero-length line: displays as a dot whose size follows its weight."""
    model_ref, dgn_model = active_model()
    eeh = EditElementHandle()
    point = DPoint3d(x, y, 0.0)
    if BentleyStatus.eSUCCESS != LineHandler.CreateLineElement(
            eeh, None, DSegment3d(point, point), dgn_model.Is3d(), model_ref):
        raise RuntimeError("LineHandler could not create an element.")
    if symbology is not None:
        _apply_symbology(eeh, symbology)
    return eeh


def zoom_to_extent(extent):
    view_set = IViewManager.GetActiveViewSet()
    viewport = view_set.GetSelectedViewport()
    if viewport is None:
        return
    x_min, y_min, x_max, y_max = extent
    delta = viewport.GetViewDelta()
    if delta.x <= 0.0 or delta.y <= 0.0:
        return
    factor = 1.2 * max((x_max - x_min) / delta.x, (y_max - y_min) / delta.y)
    if factor <= 0.0:
        return
    centre = DPoint3d((x_min + x_max) / 2.0, (y_min + y_max) / 2.0, 0.0)
    viewport.Zoom(centre, factor, True)
    viewport.SynchWithViewInfo(True, True)
    view_set.UpdateView(viewport, DgnDrawMode.eDRAW_MODE_Normal, DrawPurpose.eUpdate,
                        IndexedViewSet.FullUpdateInfo())


def _item_type_properties(event):
    string_type = CustomProperty.Type1.eString
    double_type = CustomProperty.Type1.eDouble
    info = event.route_info
    from_lat, from_lon = event.lat_lon(event.from_measure)
    end_measure = event.to_measure if event.event_type == EVENT_LINEAR else event.from_measure
    to_lat, to_lon = event.lat_lon(end_measure)
    return [
        ("Route", string_type, info.route),
        ("Name", string_type, info.name_id),
        ("Jurisdiction", string_type, info.jurisdiction),
        ("Category", string_type, info.category),
        ("EventType", string_type, event.event_type),
        ("StartMeasure", double_type, round(event.from_measure, 3) + 0.0),
        ("EndMeasure", double_type, round(end_measure, 3) + 0.0),
        ("EventLength", double_type, round(event.length, 3)),
        ("RouteLength", double_type, round(info.geometry.m_max, 3)),
        ("FromLatitude", double_type, round(from_lat, 6)),
        ("FromLongitude", double_type, round(from_lon, 6)),
        ("ToLatitude", double_type, round(to_lat, 6)),
        ("ToLongitude", double_type, round(to_lon, 6)),
    ]


def _item_type_library(dgn_file, properties):
    """Return the item type library, creating or extending it as needed.

    The caller must keep the returned library referenced for as long as it uses
    item types from it: an ItemType is owned by its library, and using one after
    the library object has been released crashes MicroStation.
    """
    library = ItemTypeLibrary.FindByName(ITEM_LIBRARY, dgn_file)
    changed = False
    if library is None:
        library = ItemTypeLibrary(ITEM_LIBRARY, dgn_file, False)
        changed = True
    graphic_properties = [(GRAPHIC_PART, CustomProperty.Type1.eString, "")]
    for type_name, type_properties in ((ITEM_TYPE, properties),
                                       (GRAPHIC_TYPE, graphic_properties)):
        item_type = library.GetItemTypeByName(type_name)
        if item_type is None:
            item_type = library.AddItemType(type_name, False)
            changed = True
        for prop_name, prop_type, _value in type_properties:
            if item_type.GetPropertyByName(prop_name) is None:
                item_type.AddProperty(prop_name, False).SetType(prop_type)
                changed = True
    if changed:
        if not library.Write():
            raise RuntimeError(f"Could not write item type library '{ITEM_LIBRARY}'.")
        library = ItemTypeLibrary.FindByName(ITEM_LIBRARY, dgn_file)
    for type_name in (ITEM_TYPE, GRAPHIC_TYPE):
        if library is None or library.GetItemTypeByName(type_name) is None:
            raise RuntimeError(f"Item type '{ITEM_LIBRARY} / {type_name}' is not available.")
    return library


def _attach_metadata(eeh, library, properties):
    item_type = library.GetItemTypeByName(ITEM_TYPE)
    host = CustomItemHost(eeh, False)      # kept referenced until the instance is written
    instance = host.ApplyCustomItem(item_type)
    if instance is None:
        raise RuntimeError("The item type could not be applied to the element.")
    for prop_name, _prop_type, value in properties:
        instance.SetValue(prop_name, ECValue(value))
    instance.WriteChanges()


def _tag_graphic(eeh, library, part):
    """Mark an element the tool added (label, marker or sign) so Delete can find it."""
    item_type = library.GetItemTypeByName(GRAPHIC_TYPE)
    host = CustomItemHost(eeh, False)
    instance = host.ApplyCustomItem(item_type)
    if instance is None:
        raise RuntimeError("The LRS Graphic item type could not be applied to the element.")
    instance.SetValue(GRAPHIC_PART, ECValue(part))
    instance.WriteChanges()


def _created_kind(eh):
    """ITEM_TYPE or the Part of an element the tool created, else None."""
    host = CustomItemHost(eh)
    if host.GetCustomItem(ITEM_LIBRARY, ITEM_TYPE) is not None:
        return ITEM_TYPE
    instance = host.GetCustomItem(ITEM_LIBRARY, GRAPHIC_TYPE)
    if instance is None:
        return None
    value = ECValue()
    instance.GetValue(value, GRAPHIC_PART)
    return value.GetString() or GRAPHIC_TYPE


def find_created_elements():
    """Element refs in the active model that the tool created, with a count per kind."""
    _model_ref, dgn_model = active_model()
    refs, counts = [], {}
    for ref in dgn_model.GetGraphicElements():
        eh = ElementHandle(ref)
        if not eh.IsValid():
            continue
        kind = _created_kind(eh)
        if kind is not None:
            refs.append(ref)
            counts[kind] = counts.get(kind, 0) + 1
    return refs, counts


def delete_created_elements(refs):
    """Delete the given elements as one undoable step. Returns the number deleted."""
    model_ref, _dgn_model = active_model()
    txns = ITxnManager.GetManager()
    txns.CloseCurrentTxn(False)            # keep earlier work out of this undo step
    deleted = 0
    for ref in refs:
        eeh = EditElementHandle(ref, model_ref)
        if eeh.IsValid() and eeh.DeleteFromModel() == 0:
            deleted += 1
    txns.CloseCurrentTxn(False)
    return deleted


def _place_label(dgn_model, text, x, y, height, library):
    """Place text centred on (x, y)."""
    run_props = RunProperties.Create(
        DgnFontManager.GetDefaultTrueTypeFont(), DPoint2d(height, height), dgn_model)
    block_props = TextBlockProperties.Create(dgn_model)
    paragraph = ParagraphProperties.Create(dgn_model)
    paragraph.SetJustification(TextElementJustification.eCenterMiddle)
    block = TextBlock(block_props, paragraph, run_props, dgn_model)
    block.AppendText(text)
    block.SetUserOrigin(DPoint3d(x, y, 0.0))
    eeh = EditElementHandle()
    if (TextHandlerBase.CreateElement(eeh, None, block)
            != TextBlockToElementResult.eTEXTBLOCK_TO_ELEMENT_RESULT_Success):
        raise RuntimeError(f"Could not create label '{text}'.")
    ElementPropertyUtils.ApplyActiveSettings(eeh)
    if eeh.AddToModel() != 0:
        raise RuntimeError(f"Could not add label '{text}' to the model.")
    _tag_graphic(eeh, library, PART_LABEL)


def _colour_id(rgb):
    return DgnColorMap.CreateElementColor(
        IntColorDef(rgb[0], rgb[1], rgb[2]), None, None, ISessionMgr.GetActiveDgnFile())


def _milepost_child(eeh, colour, weight=0, fill=None):
    """Put a sign component on the active level with its own colour (and solid fill)."""
    ElementPropertyUtils.ApplyActiveSettings(eeh)
    setter = ElementPropertiesSetter()
    setter.SetColor(colour)
    setter.SetWeight(weight)
    if fill is not None:
        setter.SetFillColor(fill)
    setter.Apply(eeh)
    if fill is not None:
        eeh.GetHandler().AddSolidFill(eeh, fill, True)
    return eeh


def milepost_cell(x, y, digits, decimal, unit, green, white, centre):
    """A milepost sign as one cell: green panel, white border and legend, on a post.

    The sign stands upright with its middle at `centre`; the post runs from the route
    point (x, y) to the sign.
    """
    model_ref, dgn_model = active_model()
    is_3d = dgn_model.Is3d()
    width, height, texts, rule = milepost_layout(digits, decimal, unit)
    route_x, route_y = x, y
    x, centre_y = centre
    bottom = centre_y - height / 2.0
    children = []

    post = EditElementHandle()
    LineHandler.CreateLineElement(
        post, None, DSegment3d(DPoint3d(route_x, route_y, 0.0), DPoint3d(x, centre_y, 0.0)),
        is_3d, model_ref)
    children.append(_milepost_child(post, green, 2))

    border = 0.1 * unit
    for inset, colour in ((0.0, white), (border, green)):
        outline = rounded_rectangle(x, bottom + inset, width - 2 * inset, height - 2 * inset,
                                    0.4 * unit - inset)
        panel = EditElementHandle()
        if BentleyStatus.eSUCCESS != ShapeHandler.CreateShapeElement(
                panel, None, DPoint3dArray([DPoint3d(px, py, 0.0) for px, py in outline]),
                is_3d, model_ref):
            raise RuntimeError("ShapeHandler could not create the milepost panel.")
        children.append(_milepost_child(panel, colour, 0, colour))

    if rule is not None:
        line = EditElementHandle()
        LineHandler.CreateLineElement(
            line, None, DSegment3d(DPoint3d(x - 0.3 * width, bottom + rule, 0.0),
                                   DPoint3d(x + 0.3 * width, bottom + rule, 0.0)),
            is_3d, model_ref)
        children.append(_milepost_child(line, white, 2))

    font = DgnFontManager.GetDefaultTrueTypeFont()
    for text, centre_y, text_height in texts:
        run_props = RunProperties.Create(
            font, DPoint2d(0.85 * text_height, text_height), dgn_model)
        run_props.SetIsBold(True)
        run_props.SetColor(white)
        paragraph = ParagraphProperties.Create(dgn_model)
        paragraph.SetJustification(TextElementJustification.eCenterMiddle)
        block_props = TextBlockProperties.Create(dgn_model)
        block = TextBlock(block_props, paragraph, run_props, dgn_model)
        block.AppendText(text)
        block.SetUserOrigin(DPoint3d(x, bottom + centre_y, 0.0))
        legend = EditElementHandle()
        if (TextHandlerBase.CreateElement(legend, None, block)
                != TextBlockToElementResult.eTEXTBLOCK_TO_ELEMENT_RESULT_Success):
            raise RuntimeError(f"Could not create the milepost legend '{text}'.")
        children.append(_milepost_child(legend, white))

    cell = EditElementHandle()
    NormalCellHeaderHandler.CreateOrphanCellElement(cell, "MILEPOST", is_3d, model_ref)
    for child in children:
        NormalCellHeaderHandler.AddChildElement(cell, child)
    NormalCellHeaderHandler.AddChildComplete(cell)
    return cell


def place_mileposts(event, geometry_uor, mode, length_uors, log, library):
    """Add a milepost sign cell at each measure of the chosen mode. Returns the count."""
    measures = event.milepost_measures(mode)
    if not measures:
        return 0
    if event.event_type == EVENT_LINEAR and mode != MILEPOST_ENDS:
        sign_height = 0.35 * mode * length_uors / event.length
    elif length_uors > 0.0:
        sign_height = length_uors / 15.0
    else:
        x_min, y_min, x_max, y_max = geometry_uor.extent()
        sign_height = math.hypot(x_max - x_min, y_max - y_min) / 30.0
    uor_per_meter = active_model()[1].GetModelInfo().GetUorPerMeter()
    sign_height = clamp(sign_height, SIGN_HEIGHT_M[0] * uor_per_meter,
                        SIGN_HEIGHT_M[1] * uor_per_meter)
    unit = MILEPOST_SIZE * sign_height / 5.8      # 5.8 units = two digits plus a decimal
    green, white = _colour_id(MILEPOST_GREEN), _colour_id(MILEPOST_WHITE)
    for measure in measures:
        x, y = geometry_uor.point_at(measure)
        digits, decimal = milepost_text(measure, mode)
        width, height, _texts, _rule = milepost_layout(digits, decimal, unit)
        centre = beside_route(geometry_uor, measure, 1, width, height, 0.8 * unit)
        cell = milepost_cell(x, y, digits, decimal, unit, green, white, centre)
        if cell.AddToModel() != 0:
            raise RuntimeError(f"Milepost {fmt_measure(measure)} could not be added.")
        _tag_graphic(cell, library, PART_MILEPOST)
    log(f"Placed {len(measures)} milepost sign(s) to the right of the direction of travel.")
    return len(measures)


def create_event_elements(event, geometry_uor, labels, mileposts=MILEPOST_NONE, log=print):
    """Add the event to the active model. Returns the new element ids."""
    event.validate()
    sign_count = len(event.milepost_measures(mileposts))
    if sign_count > MAX_MILEPOSTS:
        raise ValueError(f"{sign_count} milepost signs is more than the limit of "
                         f"{MAX_MILEPOSTS}. Choose a wider spacing or a shorter event.")
    model_ref, dgn_model = active_model()
    is_3d = dgn_model.Is3d()
    if not is_3d:
        log("Active model is 2D - elements are created 2D.")

    # (element, symbology override or None, True for the LRS metadata or the marker Part)
    from_xy = geometry_uor.point_at(event.from_measure)
    if event.event_type == EVENT_POINT:
        elements = [(point_element(*from_xy), POINT_EVENT, True)]
        to_xy, length_uors = from_xy, 0.0
    else:
        parts = geometry_uor.slice(event.from_measure, event.to_measure)
        if not parts:
            raise RuntimeError("The route has no geometry between these measures.")
        to_xy = geometry_uor.point_at(event.to_measure)
        length_uors = sum(polyline_length(part) for part in parts)
        elements = []
        for part in parts:
            chunks = linestring_elements(part)
            if len(chunks) == 1:
                elements.append((chunks[0], None, True))
                continue
            chain = EditElementHandle()
            ChainHeaderHandler.CreateChainHeaderElement(chain, None, False, is_3d, model_ref)
            for chunk in chunks:
                ChainHeaderHandler.AddComponentElement(chain, chunk)
            ChainHeaderHandler.AddComponentComplete(chain)
            elements.append((chain, None, True))
        elements.append((point_element(*from_xy), CREATED_FROM, PART_FROM))
        elements.append((point_element(*to_xy), CREATED_TO, PART_TO))

    properties = _item_type_properties(event)
    library = _item_type_library(ISessionMgr.GetActiveDgnFile(), properties)
    element_ids = []
    for eeh, symbology, has_metadata in elements:
        ElementPropertyUtils.ApplyActiveSettings(eeh)
        if symbology is not None:
            _apply_symbology(eeh, symbology)
        if eeh.AddToModel() != 0:
            raise RuntimeError("An element could not be added to the model.")
        if has_metadata is True:
            _attach_metadata(eeh, library, properties)
            element_ids.append(eeh.GetElementId())
        else:
            _tag_graphic(eeh, library, has_metadata)

    info = event.route_info
    log(f"Added {len(element_ids)} event element(s) on the active level: "
        f"{', '.join(str(i) for i in element_ids)}.")
    if event.event_type == EVENT_LINEAR:
        log("Added the start marker (green) and stop marker (red).")
        # Reported in metres: GetUorPerMeter is on every platform version this script
        # supports (MicroStation 2026 and OpenRoads Designer 2025).
        metres = length_uors / dgn_model.GetModelInfo().GetUorPerMeter()
        log(f"Geometry length {metres:,.1f} m ({metres / 1609.344:.3f} mi) "
            f"for a {event.length:.3f} mi event.")
    else:
        log("The point event is drawn yellow.")
    log(f"Attached '{ITEM_LIBRARY} / {ITEM_TYPE}': Route={info.route}, Name={info.name_id}, "
        f"StartMeasure={fmt_measure(event.from_measure)}, "
        f"EndMeasure={fmt_measure(properties[6][2])}.")

    if labels != LABELS_NONE:
        x_min, y_min, x_max, y_max = geometry_uor.extent()
        uor_per_meter = dgn_model.GetModelInfo().GetUorPerMeter()
        height = clamp(
            (length_uors / 75.0) or math.hypot(x_max - x_min, y_max - y_min) / 150.0,
            LABEL_HEIGHT_M[0] * uor_per_meter, LABEL_HEIGHT_M[1] * uor_per_meter)
        # (text, measure it belongs to, lines stacked beyond the first label there)
        texts = [(f"MP {fmt_measure(event.from_measure)}", event.from_measure, 0)]
        if event.event_type == EVENT_LINEAR:
            texts.append((f"MP {fmt_measure(event.to_measure)}", event.to_measure, 0))
        if labels == LABELS_ALL:
            if event.event_type == EVENT_LINEAR:
                texts.append((f"{info.route}  {info.name_id}  ({event.length:.3f} mi)",
                              (event.from_measure + event.to_measure) / 2.0, 0))
            else:
                texts.append((f"{info.route}  {info.name_id}", event.from_measure, 1))
        try:
            for text, measure, line in texts:
                # A second line is stacked on the first, on the side away from the route.
                upwards = 1.0 if geometry_uor.direction_at(measure)[0] >= 0.0 else -1.0
                centre_x, centre_y = beside_route(
                    geometry_uor, measure, -1, LABEL_CHAR_WIDTH * height * len(text), height,
                    0.8 * height, line * 1.6 * height * upwards)
                _place_label(dgn_model, text, centre_x, centre_y, height, library)
            log("Placed labels to the left of the direction of travel.")
        except Exception as ex:
            log(f"WARNING: labels were not placed ({ex}).")

    if mileposts != MILEPOST_NONE:
        try:
            place_mileposts(event, geometry_uor, mileposts, length_uors, log, library)
        except Exception as ex:
            log(f"WARNING: milepost signs were not all placed ({ex}).")
            print(traceback.format_exc())

    MessageCenter.ShowInfoMessage(
        f"LRS event {info.route} created ({len(element_ids)} element(s)).",
        f"{info.name_id}  measure {fmt_measure(event.from_measure)} to "
        f"{fmt_measure(properties[6][2])} mi", False)
    return element_ids


if IN_MSTN:

    class LrsPickTool(DgnPrimitiveTool):
        """Previews the route/event as dynamics and picks measures with data points."""

        active = None

        def __init__(self, dialog):
            DgnPrimitiveTool.__init__(self, 0, 0)  # C++ base's __init__ must be called.
            self.m_self = self                     # Keep the tool alive while installed.
            self.dialog = dialog
            self.reported_error = False
            self.last_prompt = ""

        def _OnPostInstall(self):
            DgnPrimitiveTool._OnPostInstall(self)
            self._BeginDynamics()
            self._prompt(self.dialog.pick_prompt())

        def _OnRestartTool(self):
            LrsPickTool.install(self.dialog)

        def _OnModelEndDrag(self, ev):
            return False

        def _prompt(self, text):
            if text != self.last_prompt:
                self.last_prompt = text
                NotificationManager.OutputPrompt(text)

        def _OnDataButton(self, ev):
            try:
                point = ev.GetPoint()
                self.dialog.on_pick(point.x, point.y)
                self._prompt(self.dialog.pick_prompt())
            except Exception as ex:
                self.dialog.log(f"ERROR while picking: {ex}")
                print(traceback.format_exc())
            return False

        def _OnResetButton(self, ev):
            self._ExitTool()
            return True

        def _OnDynamicFrame(self, ev):
            try:
                point = ev.GetPoint()
                elements, measure = self.dialog.dynamic_elements(point.x, point.y)
                redraw = RedrawElems()
                redraw.SetDynamicsViews(IViewManager.GetActiveViewSet(), ev.GetViewport())
                redraw.SetDrawMode(eDRAW_MODE_TempDraw)
                redraw.SetDrawPurpose(DrawPurpose.eDynamics)
                for eeh in elements:
                    redraw.DoRedraw(eeh)
                if measure is not None:
                    self._prompt(f"{self.dialog.pick_prompt()}  MP {fmt_measure(measure)}")
            except Exception as ex:
                if not self.reported_error:
                    self.reported_error = True
                    self.dialog.log(f"ERROR in the drawing preview: {ex}")
                    print(traceback.format_exc())

        def _OnCleanup(self):
            if LrsPickTool.active is self:
                LrsPickTool.active = None
                self.dialog.on_tool_ended()
            self.m_self = None

        @staticmethod
        def install(dialog):
            tool = LrsPickTool(dialog)
            LrsPickTool.active = tool
            tool.InstallTool()


# ---------------------------------------------------------------------------
# User interface
# ---------------------------------------------------------------------------
# VDOT logo (154 x 42 PNG from vdot.virginia.gov), base64-encoded so the script needs
# no separate image file.
VDOT_LOGO_PNG = (
    "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"
)
THEMES = {
    "Dark": {
        "BG": "#0f172a", "TEXT": "#e2e8f0", "CARD": "#1e293b", "CARD_BORDER": "#2c3a52",
        "TITLE": "#f8fafc", "MUTED": "#94a3b8", "VALUE": "#cbd5e1",
        "ACCENT": "#f97316", "ACCENT_HOVER": "#fb923c", "ACCENT_PRESSED": "#ea580c",
        "ON_ACCENT": "#0f172a", "PIN_FROM": "#4ade80", "PIN_TO": "#f87171",
        "INPUT_BG": "#0f172a", "INPUT_BORDER": "#334155",
        "DISABLED_TEXT": "#64748b", "DISABLED_BG": "#1e293b", "DISABLED_BORDER": "#243044",
        "LOG_BG": "#0b1220", "GHOST_TEXT": "#cbd5e1", "GHOST_HOVER": "#64748b",
        "LOGO_BG": "#ffffff", "CARD_TITLE": "#f97316",
        "PICK": "#22d3ee", "PICK_BG": "#0e3a4a",
        "PIN_FROM_BG": "#14532d", "PIN_TO_BG": "#7f1d1d",
        "LOG_ERROR": "#f87171", "LOG_WARN": "#fbbf24", "LOG_OK": "#4ade80",
    },
    "Light": {
        "BG": "#f1f5f9", "TEXT": "#0f172a", "CARD": "#ffffff", "CARD_BORDER": "#cbd5e1",
        "TITLE": "#003a70", "MUTED": "#475569", "VALUE": "#1e293b",
        "ACCENT": "#ea580c", "ACCENT_HOVER": "#f97316", "ACCENT_PRESSED": "#c2410c",
        "ON_ACCENT": "#ffffff", "PIN_FROM": "#16a34a", "PIN_TO": "#dc2626",
        "INPUT_BG": "#f8fafc", "INPUT_BORDER": "#cbd5e1",
        "DISABLED_TEXT": "#94a3b8", "DISABLED_BG": "#f1f5f9", "DISABLED_BORDER": "#e2e8f0",
        "LOG_BG": "#f8fafc", "GHOST_TEXT": "#334155", "GHOST_HOVER": "#64748b",
        "LOGO_BG": "#ffffff", "CARD_TITLE": "#ea580c",
        "PICK": "#0891b2", "PICK_BG": "#ecfeff",
        "PIN_FROM_BG": "#dcfce7", "PIN_TO_BG": "#fee2e2",
        "LOG_ERROR": "#dc2626", "LOG_WARN": "#b45309", "LOG_OK": "#15803d",
    },
}

# Status badge colours: (background, text).
STATUS_BUSY = ("#22d3ee", "#0f172a")
STATUS_OK = ("#4ade80", "#052e16")
STATUS_OFFLINE = ("#dc2626", "#ffffff")

STYLE = """
QWidget#root { background: @BG@; }
QScrollArea#cards, QWidget#body { background: transparent; border: none; }
QScrollBar:vertical { background: transparent; width: 10px; margin: 0; }
QScrollBar::handle:vertical { background: @INPUT_BORDER@; border-radius: 5px; min-height: 30px; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }
QWidget { color: @TEXT@; font-family: 'Segoe UI'; font-size: 10pt; }
QFrame#card { background: @CARD@; border: 1px solid @CARD_BORDER@; border-radius: 12px; }
QLabel#title { font-size: 15pt; font-weight: 600; color: @TITLE@; }
QLabel#subtitle { color: @MUTED@; font-size: 9pt; }
QLabel#logo { background: @LOGO_BG@; border-radius: 8px; padding: 5px 8px; }
QLabel#logoText { background: @LOGO_BG@; color: #003a70; border-radius: 8px;
                  padding: 5px 10px; font-size: 14pt; font-weight: 800; }
QLabel#cardTitle { color: @CARD_TITLE@; font-size: 8pt; font-weight: 700; letter-spacing: 1px; }
QLabel#field { color: @MUTED@; font-size: 9pt; }
QLabel#hint { color: @MUTED@; font-size: 9pt; }
QLabel#hint[error="true"] { color: @LOG_ERROR@; }
QLabel#value { color: @VALUE@; font-family: 'Cascadia Mono', 'Consolas'; font-size: 9pt; }
QLabel#pinFrom { color: @PIN_FROM@; font-weight: 600; border-radius: 8px; padding: 3px 6px; }
QLabel#pinTo { color: @PIN_TO@; font-weight: 600; border-radius: 8px; padding: 3px 6px; }
QLabel#pinFrom[next="true"] { background: @PIN_FROM_BG@; }
QLabel#pinTo[next="true"] { background: @PIN_TO_BG@; }
QComboBox { background: @INPUT_BG@; border: 1px solid @INPUT_BORDER@; border-radius: 7px;
            padding: 6px 9px; selection-background-color: @ACCENT@; min-height: 18px; }
QComboBox:focus, QComboBox:on { border-color: @ACCENT@; }
QComboBox:disabled { color: @DISABLED_TEXT@; border-color: @DISABLED_BORDER@; }
QComboBox::drop-down { subcontrol-origin: padding; subcontrol-position: center right;
    border: none; border-left: 1px solid @INPUT_BORDER@; width: 26px; }
QComboBox::down-arrow { image: url(@ARROW@); width: 10px; height: 6px; }
QComboBox::down-arrow:disabled { image: url(@ARROW_DISABLED@); }
QComboBox QLineEdit { background: transparent; border: none; padding: 0;
                      selection-background-color: @ACCENT@; }
QComboBox QAbstractItemView, QAbstractItemView { background: @CARD@; color: @TEXT@;
    border: 1px solid @INPUT_BORDER@; selection-background-color: @ACCENT@;
    selection-color: @ON_ACCENT@; outline: none; }
QPushButton { border-radius: 9px; padding: 9px 16px; }
QPushButton#primary { background: @ACCENT@; color: @ON_ACCENT@; font-weight: 700; border: none; }
QPushButton#primary:hover { background: @ACCENT_HOVER@; }
QPushButton#primary:pressed { background: @ACCENT_PRESSED@; }
QPushButton#ghost, QPushButton#mini { background: transparent; color: @GHOST_TEXT@;
                                      border: 1px solid @INPUT_BORDER@; }
QPushButton#mini { padding: 3px 10px; border-radius: 7px; font-size: 9pt; }
QPushButton#ghost:hover, QPushButton#mini:hover { border-color: @GHOST_HOVER@; color: @TITLE@; }
QPushButton#ghost:checked { border: 1px solid @PICK@; background: @PICK_BG@; color: @TITLE@;
                            font-weight: 600; }
QPushButton#segLeft, QPushButton#segRight { background: @INPUT_BG@; color: @GHOST_TEXT@;
    border: 1px solid @INPUT_BORDER@; padding: 6px 14px; border-radius: 7px;
    min-width: 54px; }
QPushButton#segLeft { border-top-right-radius: 0; border-bottom-right-radius: 0; }
QPushButton#segRight { border-top-left-radius: 0; border-bottom-left-radius: 0;
                       border-left: none; }
QPushButton#segLeft:checked, QPushButton#segRight:checked { background: @ACCENT@;
    color: @ON_ACCENT@; border-color: @ACCENT@; font-weight: 700; }
QPushButton#logToggle { background: transparent; border: none; color: @MUTED@;
    padding: 0 2px; font-size: 8pt; font-weight: 700; text-align: left; }
QPushButton#logToggle:hover { color: @TITLE@; }
QPushButton#danger { background: transparent; color: @LOG_ERROR@;
                     border: 1px solid @LOG_ERROR@; }
QPushButton#danger:hover { background: @LOG_ERROR@; color: #ffffff; }
QPushButton#danger:disabled {
    color: @DISABLED_TEXT@; background: @DISABLED_BG@; border: 1px solid @DISABLED_BORDER@; }
QPushButton#primary:disabled, QPushButton#ghost:disabled, QPushButton#mini:disabled {
    color: @DISABLED_TEXT@; background: @DISABLED_BG@; border: 1px solid @DISABLED_BORDER@; }
QFrame#offer { background: @PICK_BG@; border: 1px solid @PICK@; border-radius: 10px; }
QLabel#offerText { background: transparent; color: @TITLE@; font-weight: 600; }
QMessageBox { background: @CARD@; }
QMessageBox QPushButton { background: @INPUT_BG@; color: @TEXT@; border: 1px solid @INPUT_BORDER@;
    padding: 6px 18px; min-width: 64px; }
QMessageBox QPushButton:hover { border-color: @ACCENT@; }
QMessageBox QPushButton:default { background: @ACCENT@; color: @ON_ACCENT@; border-color: @ACCENT@;
    font-weight: 700; }
QTextEdit { background: @LOG_BG@; border: 1px solid @DISABLED_BORDER@; border-radius: 9px;
    font-family: 'Cascadia Mono', 'Consolas'; font-size: 9pt; color: @VALUE@; padding: 6px; }
"""

if HAVE_QT:

    _ARROW_IMAGES = {}

    def _arrow_image(colour):
        """Drop-down arrow as a temp PNG (Qt style sheets only take image files).

        Each colour is drawn and written once per session.
        """
        if colour not in _ARROW_IMAGES:
            pixmap = QPixmap(20, 12)
            pixmap.fill(Qt.transparent)
            painter = QPainter(pixmap)
            painter.setRenderHint(QPainter.Antialiasing)
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor(colour))
            painter.drawPolygon(QPolygonF([QPointF(0, 0), QPointF(20, 0), QPointF(10, 12)]))
            painter.end()
            path = (tempfile.gettempdir().replace("\\", "/")
                    + f"/vdot_lrs_arrow_{colour.lstrip('#')}.png")
            pixmap.save(path)
            _ARROW_IMAGES[colour] = path
        return _ARROW_IMAGES[colour]

    def build_style(theme):
        palette = THEMES[theme]
        style = STYLE
        for token, colour in palette.items():
            style = style.replace(f"@{token}@", colour)
        return (style.replace("@ARROW@", _arrow_image(palette["ACCENT"]))
                .replace("@ARROW_DISABLED@", _arrow_image(palette["DISABLED_TEXT"])))

    class EventDialog(QWidget):
        def __init__(self, service=None, settings=None):
            super().__init__()
            self.service = service or LrsService()
            self.route_info = None
            self.geometry_uor = None
            self.route_preview = []
            self.event_preview = []
            self.next_pick = "from"
            self.settings = settings or QSettings("VDOT LRS", "Event to Alignment")
            self._last_event_key = None
            self._log_lines = []

            # Service requests run on a worker thread and hand their result back through
            # this queue; only the newest request's result is used.
            self._results = queue.Queue()
            self._request = 0
            self._on_result = None
            self._worker = None
            self._busy_text = None
            self._busy_step = 0
            self.service.on_retry = lambda attempt: self._results.put(("retry", 0, attempt))

            self.setObjectName("root")
            self.setWindowTitle(f"{APP_TITLE}  \u2014  Created by {APP_AUTHOR}  \u2014  "
                                f"Build {BUILD_DATE}")
            self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
            screen = QApplication.primaryScreen()
            available = screen.availableGeometry().height() if screen is not None else 900
            self.resize(520, max(480, min(900, available - 70)))

            self.theme = self._combo()
            self.theme.addItems(list(THEMES))
            saved_theme = self.settings.value("theme", "Dark")
            self.theme.setCurrentText(saved_theme if saved_theme in THEMES else "Dark")
            self.theme.activated.connect(lambda _i: self._apply_theme())

            self.jurisdiction = self._combo(searchable=True)
            self.route = self._combo(searchable=True)
            self.linear_button = self._button(EVENT_LINEAR, "segLeft", checkable=True)
            self.point_button = self._button(EVENT_POINT, "segRight", checkable=True)
            self.linear_button.setChecked(True)
            self.event_type = QButtonGroup(self)
            self.event_type.addButton(self.linear_button)
            self.event_type.addButton(self.point_button)
            self.from_measure = self._combo(editable=True)
            self.to_measure = self._combo(editable=True)
            self.coordinates = self._combo()
            self.labels = self._combo()
            self.labels.addItem("Measures and route name", LABELS_ALL)
            self.labels.addItem("Measures only", LABELS_MEASURES)
            self.labels.addItem("No labels", LABELS_NONE)
            self.mileposts = self._combo()
            for text, mode in MILEPOST_CHOICES:
                self.mileposts.addItem(text, mode)
            self._restore_choice(self.labels, "labels")
            self._restore_choice(self.mileposts, "mileposts")

            self.route_hint = self._label("", "hint")
            self.route_hint.setWordWrap(True)
            self.from_label = self._label("", "pinFrom")
            self.to_label = self._label("", "pinTo")
            for label in (self.from_label, self.to_label):
                label.setMinimumWidth(84)
            self.from_location = self._label("", "value")
            self.to_location = self._label("", "value")
            self.event_hint = self._label("", "hint")
            self.event_hint.setWordWrap(True)
            self.model_hint = self._label("", "hint")
            self.model_hint.setWordWrap(True)
            self.status = self._label("", "status")
            self.log_view = QTextEdit()
            self.log_view.setReadOnly(True)
            self.log_view.setFixedHeight(96)
            self.log_toggle = self._button("", "logToggle")
            self.log_toggle.setCursor(Qt.PointingHandCursor)
            self.opposite_button = self._button("\u21c4  Opposite direction", "mini")
            self.swap_button = self._button("\u21c5", "mini")
            self.swap_button.setToolTip("Swap the From and To measures")
            self.full_button = self._button("Full route", "mini")
            self.full_button.setToolTip("Set From and To to the start and end of the route")
            self.pick_button = self._button("", "ghost", checkable=True)
            self.zoom_button = self._button("Zoom to route", "ghost")
            self.create_button = self._button("Create alignment", "primary")
            self.delete_button = self._button("Delete created", "danger")
            self.delete_button.setToolTip(
                "Delete every event, marker, label and milepost sign this tool placed "
                "in the active model")
            self.delete_button.setEnabled(IN_MSTN)
            self.offer_bar = QFrame()
            self.offer_bar.setObjectName("offer")
            self.offer_text = self._label("", "offerText")
            self.offer_text.setWordWrap(True)
            self.offer_yes = self._button("Yes, draw it", "primary")
            self.offer_no = self._button("No", "ghost")
            self.offer_bar.hide()
            self._build_layout()
            self._apply_theme()
            self._set_status("CONNECTING", STATUS_BUSY)
            self._set_log_open(self.settings.value("log_open", "1") != "0")
            self._load_model_state()
            self._sync_pick_state()

            self.jurisdiction.activated.connect(lambda _i: self._guard(self._load_routes))
            self.route.activated.connect(lambda _i: self._guard(self._load_route))
            self.event_type.buttonClicked.connect(
                lambda _b: self._guard(self._measures_edited))
            self.coordinates.activated.connect(lambda _i: self._guard(self._coordinates_chosen))
            self.labels.activated.connect(
                lambda _i: self._save_choice(self.labels, "labels"))
            self.mileposts.activated.connect(
                lambda _i: self._save_choice(self.mileposts, "mileposts"))
            for combo in (self.from_measure, self.to_measure):
                combo.activated.connect(lambda _i: self._guard(self._measures_edited))
                combo.lineEdit().editingFinished.connect(
                    lambda: self._guard(self._measures_edited))
            self.opposite_button.clicked.connect(lambda: self._guard(self._switch_opposite))
            self.swap_button.clicked.connect(lambda: self._guard(self._swap))
            self.full_button.clicked.connect(lambda: self._guard(self._full_route))
            self.log_toggle.clicked.connect(
                lambda: self._set_log_open(not self.log_view.isVisibleTo(self)))
            self.pick_button.clicked.connect(lambda: self._guard(self._toggle_pick_tool))
            self.zoom_button.clicked.connect(lambda: self._guard(self._zoom))
            self.create_button.clicked.connect(lambda: self._guard(self._create))
            self.delete_button.clicked.connect(lambda: self._guard(self._delete_created))
            self.offer_yes.clicked.connect(lambda: self._guard(self._accept_offer))
            self.offer_no.clicked.connect(self._decline_offer)

            self._result_timer = QTimer(self)
            self._result_timer.setInterval(40)
            self._result_timer.timeout.connect(self._poll_results)
            self._result_timer.start()
            self._busy_timer = QTimer(self)
            self._busy_timer.setInterval(350)
            self._busy_timer.timeout.connect(self._animate_status)

            QTimer.singleShot(0, lambda: self._guard(self._connect))

        # -- construction ---------------------------------------------------
        @staticmethod
        def _combo(editable=False, searchable=False):
            combo = QComboBox()
            combo.setMaxVisibleItems(18)
            # Without this the popup ignores maxVisibleItems and fills the screen.
            combo.setStyleSheet("QComboBox { combobox-popup: 0; }")
            # Lets a drop-down shrink below its longest entry, so the cards always fit
            # the dialog width (also when the scroll bar is showing).
            combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
            combo.setMinimumContentsLength(6)
            if editable or searchable:
                combo.setEditable(True)
                combo.setInsertPolicy(QComboBox.NoInsert)
            if searchable:
                # Typing filters the list to entries containing the text, in any case.
                completer = combo.completer()
                completer.setCompletionMode(QCompleter.PopupCompletion)
                completer.setFilterMode(Qt.MatchContains)
                completer.setCaseSensitivity(Qt.CaseInsensitive)
                completer.setMaxVisibleItems(14)
                combo.lineEdit().setPlaceholderText("Select or type to search")
                combo.lineEdit().editingFinished.connect(
                    lambda: EventDialog._show_selection(combo))
            elif editable:
                combo.setCompleter(None)
            return combo

        @staticmethod
        def _show_selection(combo):
            """After a search, show the selected entry again rather than leftover typing."""
            index = combo.currentIndex()
            if index >= 0:
                combo.setEditText(combo.itemText(index))
            else:
                combo.clearEditText()

        @staticmethod
        def _label(text, object_name="field"):
            label = QLabel(text)
            label.setObjectName(object_name)
            return label

        @staticmethod
        def _button(text, object_name, checkable=False):
            button = QPushButton(text)
            button.setObjectName(object_name)
            button.setCheckable(checkable)
            return button

        @staticmethod
        def _set_flag(widget, name, on):
            """Set a style-sheet property and restyle the widget when it changes."""
            if widget.property(name) != on:
                widget.setProperty(name, on)
                widget.style().unpolish(widget)
                widget.style().polish(widget)

        def _logo(self):
            """VDOT logo from the embedded image; a text badge if it cannot be decoded."""
            pixmap = QPixmap()
            try:
                pixmap.loadFromData(base64.b64decode(VDOT_LOGO_PNG), "PNG")
            except ValueError:
                pass
            if pixmap.isNull():
                return self._label("VDOT", "logoText")
            label = self._label("", "logo")
            label.setPixmap(pixmap.scaledToHeight(34, Qt.SmoothTransformation))
            return label

        def _apply_theme(self):
            theme = self.theme.currentText()
            self.setStyleSheet(build_style(theme))
            self.settings.setValue("theme", theme)
            # Search suggestion lists are separate popup windows, so they need the theme too.
            palette = THEMES[theme]
            popup_style = (
                f"QAbstractItemView {{ background: {palette['CARD']}; color: {palette['TEXT']}; "
                f"border: 1px solid {palette['INPUT_BORDER']}; outline: none; "
                f"font-family: 'Segoe UI'; font-size: 10pt; "
                f"selection-background-color: {palette['ACCENT']}; "
                f"selection-color: {palette['ON_ACCENT']}; }}")
            for combo in (self.jurisdiction, self.route):
                combo.completer().popup().setStyleSheet(popup_style)
            # The log colours belong to the theme, so its lines are drawn again.
            self.log_view.clear()
            for message in self._log_lines:
                self.log_view.append(self._log_html(message))

        def _card(self, title):
            card = QFrame()
            card.setObjectName("card")
            layout = QVBoxLayout(card)
            layout.setContentsMargins(16, 12, 16, 14)
            layout.setSpacing(8)
            layout.addWidget(self._label(title.upper(), "cardTitle"))
            return card, layout

        def _build_layout(self):
            root = QVBoxLayout(self)
            root.setContentsMargins(18, 16, 18, 16)
            root.setSpacing(12)

            header = QHBoxLayout()
            header.setSpacing(12)
            header.addWidget(self._logo(), 0, Qt.AlignTop)
            titles = QVBoxLayout()
            titles.setSpacing(1)
            titles.addWidget(self._label("LRS Event \u2192 Alignment", "title"))
            titles.addWidget(self._label("VDOT LRS Route Master  \u00b7  MicroStation", "subtitle"))
            titles.addWidget(self._label(
                f"Created by {APP_AUTHOR}  \u00b7  Build {BUILD_DATE}", "subtitle"))
            header.addLayout(titles)
            header.addStretch(1)
            side = QVBoxLayout()
            side.setSpacing(6)
            side.addWidget(self.status, 0, Qt.AlignRight)
            side.addWidget(self.theme)
            header.addLayout(side)
            root.addLayout(header)

            # The cards sit in a scroll area so they keep their size (and scroll) when the
            # screen is short or Windows display scaling is high, instead of being squashed.
            body = QWidget()
            body.setObjectName("body")
            cards = QVBoxLayout(body)
            cards.setContentsMargins(0, 0, 0, 0)
            cards.setSpacing(12)
            scroll = QScrollArea()
            scroll.setObjectName("cards")
            scroll.setWidgetResizable(True)
            scroll.setFrameShape(QFrame.NoFrame)
            scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            scroll.setWidget(body)
            scroll.viewport().setAutoFillBackground(False)
            root.addWidget(scroll, 1)

            card, layout = self._card("Route")
            layout.addWidget(self._label("Jurisdiction"))
            layout.addWidget(self.jurisdiction)
            row = QHBoxLayout()
            row.addWidget(self._label("Route"), 0, Qt.AlignBottom)
            row.addStretch(1)
            row.addWidget(self.opposite_button)
            layout.addLayout(row)
            layout.addWidget(self.route)
            layout.addWidget(self.route_hint)
            cards.addWidget(card)

            card, layout = self._card("Event")
            grid = QGridLayout()
            grid.setHorizontalSpacing(10)
            grid.setVerticalSpacing(5)
            grid.addWidget(self._label("Type"), 0, 0)
            segment = QHBoxLayout()
            segment.setSpacing(0)
            segment.addWidget(self.linear_button)
            segment.addWidget(self.point_button)
            segment.addStretch(1)
            grid.addLayout(segment, 0, 1)
            grid.addWidget(self.full_button, 0, 2, 1, 2, Qt.AlignRight)
            grid.addWidget(self._label("Measure (mi)"), 1, 1)
            grid.addWidget(self._label("Latitude, longitude"), 1, 2)
            grid.addWidget(self.from_label, 2, 0)
            grid.addWidget(self.from_measure, 2, 1)
            grid.addWidget(self.from_location, 2, 2)
            grid.addWidget(self.to_label, 3, 0)
            grid.addWidget(self.to_measure, 3, 1)
            grid.addWidget(self.to_location, 3, 2)
            grid.addWidget(self.swap_button, 2, 3, 2, 1, Qt.AlignVCenter)
            grid.setColumnStretch(1, 2)
            grid.setColumnStretch(2, 3)
            layout.addLayout(grid)
            layout.addWidget(self.event_hint)
            cards.addWidget(card)

            card, layout = self._card("Drawing")
            grid = QGridLayout()
            grid.setHorizontalSpacing(10)
            grid.setVerticalSpacing(5)
            grid.addWidget(self._label("Coordinate system"), 0, 0, 1, 2)
            grid.addWidget(self.coordinates, 1, 0, 1, 2)
            grid.addWidget(self._label("Labels"), 2, 0)
            grid.addWidget(self._label("Milepost signs"), 2, 1)
            grid.addWidget(self.labels, 3, 0)
            grid.addWidget(self.mileposts, 3, 1)
            grid.setColumnStretch(0, 1)
            grid.setColumnStretch(1, 1)
            layout.addLayout(grid)
            layout.addWidget(self.model_hint)
            row = QHBoxLayout()
            row.setSpacing(10)
            row.addWidget(self.pick_button, 3)
            row.addWidget(self.zoom_button, 2)
            layout.addLayout(row)
            cards.addWidget(card)
            cards.addStretch(1)

            log_box = QVBoxLayout()
            log_box.setSpacing(4)
            log_box.addWidget(self.log_toggle, 0, Qt.AlignLeft)
            log_box.addWidget(self.log_view)
            root.addLayout(log_box)

            offer = QHBoxLayout(self.offer_bar)
            offer.setContentsMargins(12, 8, 8, 8)
            offer.setSpacing(8)
            offer.addWidget(self.offer_text, 1)
            offer.addWidget(self.offer_no)
            offer.addWidget(self.offer_yes)
            root.addWidget(self.offer_bar)

            buttons = QHBoxLayout()
            buttons.setSpacing(10)
            close_button = self._button("Close", "ghost")
            close_button.clicked.connect(self.close)
            buttons.addWidget(self.delete_button)
            buttons.addStretch(1)
            buttons.addWidget(close_button)
            buttons.addWidget(self.create_button)
            root.addLayout(buttons)

        def _load_model_state(self):
            gcs_name, hint = None, ""
            if IN_MSTN:
                try:
                    is_3d, gcs_name = model_summary()
                    hint = (f"Active model: {'3D' if is_3d else '2D - elements will be 2D'}"
                            f"  \u00b7  {'GCS: ' + gcs_name if gcs_name else 'no GCS assigned'}")
                except Exception as ex:
                    hint = str(ex)
            else:
                hint = "MicroStation was not detected - the drawing preview is unavailable."
            if gcs_name:
                self.coordinates.addItem(f"Model GCS  ({gcs_name})", (MODE_MODEL_GCS, 0))
            for code, name in EPSG_CHOICES:
                self.coordinates.addItem(f"EPSG:{code}  {name}", (MODE_EPSG, code))
            self.coordinates.addItem("Local tangent plane (no GCS)", (MODE_LOCAL, 0))
            if self._restore_choice(self.coordinates, "coordinates") > 0:
                self.log(f"Coordinate system from your last session: "
                         f"{self.coordinates.currentText()}.")
            self.model_hint.setText(hint)
            for button in (self.pick_button, self.zoom_button, self.create_button,
                           self.opposite_button, self.swap_button, self.full_button):
                button.setEnabled(False)

        # -- settings -------------------------------------------------------
        @staticmethod
        def _choice_key(data):
            return ":".join(str(part) for part in data) if isinstance(data, tuple) else str(data)

        def _save_choice(self, combo, key):
            self.settings.setValue(key, self._choice_key(combo.currentData()))

        def _restore_choice(self, combo, key):
            """Select the saved entry of a drop-down. Returns its index, or -1."""
            saved = self.settings.value(key)
            if saved is None:
                return -1
            for index in range(combo.count()):
                if self._choice_key(combo.itemData(index)) == str(saved):
                    combo.setCurrentIndex(index)
                    return index
            return -1

        # -- helpers --------------------------------------------------------
        def _log_html(self, message):
            palette = THEMES[self.theme.currentText()]
            colour = palette["VALUE"]
            for prefix, token in (("ERROR", "LOG_ERROR"), ("WARNING", "LOG_WARN"),
                                  ("SUCCESS", "LOG_OK")):
                if message.startswith(prefix):
                    colour = palette[token]
            return f'<span style="color: {colour};">{html.escape(message)}</span>'

        def log(self, message):
            self._log_lines.append(message)
            del self._log_lines[:-500]
            self.log_view.append(self._log_html(message))
            print(f"[LRS] {message}")

        def _set_log_open(self, is_open):
            self.log_view.setVisible(is_open)
            self.log_toggle.setText("LOG  \u25be" if is_open else "LOG  \u25b8")
            self.settings.setValue("log_open", "1" if is_open else "0")

        def _set_status(self, text, colours=STATUS_BUSY):
            self.status.setText(text)
            self.status.setStyleSheet(
                f"QLabel#status {{ background: {colours[0]}; color: {colours[1]}; "
                f"font-weight: 700; font-size: 8pt; border-radius: 9px; padding: 3px 12px; }}")

        def _set_event_hint(self, text, error=False):
            self.event_hint.setText(text)
            self._set_flag(self.event_hint, "error", error)

        def _report_error(self, ex, trace=None):
            self.log(f"ERROR: {ex}")
            if isinstance(ex, LrsOffline):
                self._set_status("OFFLINE", STATUS_OFFLINE)
                if self.route_info is None:
                    self.route_hint.setText(
                        "The service could not be reached. Choose again to retry.")
            elif not isinstance(ex, (LrsServiceError, ValueError)):
                print(trace or traceback.format_exc())
                if IN_MSTN:
                    MessageCenter.ShowErrorMessage("VDOT LRS tool error.", str(ex), False)

        def _guard(self, action):
            """Run a UI action; report failures instead of letting them escape into Qt."""
            QApplication.setOverrideCursor(Qt.WaitCursor)
            try:
                action()
            except Exception as ex:
                self._report_error(ex)
            finally:
                QApplication.restoreOverrideCursor()

        # -- background service requests -------------------------------------
        def _run_async(self, busy_text, work, done, cached=None):
            """Run work() on a worker thread, then done(result) on the main thread.

            Only the HTTP request runs on the worker. A reply already in the service's
            cache is passed to done() straight away.
            """
            if cached is not None:
                done(cached)
                return
            self._request += 1
            token = self._request
            self._on_result = done
            self._set_busy(busy_text)
            results = self._results

            def runner():
                try:
                    results.put(("done", token, work()))
                except Exception as ex:
                    results.put(("failed", token, (ex, traceback.format_exc())))

            self._worker = threading.Thread(target=runner, name="vdot-lrs-request", daemon=True)
            self._worker.start()

        def _poll_results(self):
            while True:
                try:
                    kind, token, value = self._results.get_nowait()
                except queue.Empty:
                    return
                if kind == "retry":
                    self.log(f"The VDOT LRS service did not answer - retrying ({value} of 2).")
                    continue
                if token != self._request:
                    continue                       # an older request that was superseded
                done, self._on_result = self._on_result, None
                self._set_busy(None)
                self._set_status("CONNECTED", STATUS_OK)
                if kind == "failed":
                    self._report_error(*value)
                else:
                    self._guard(lambda: done(value))

        def _set_busy(self, text):
            self._busy_text = text
            busy = text is not None
            for combo in (self.jurisdiction, self.route):
                combo.setEnabled(not busy)
            self.opposite_button.setEnabled(
                not busy and self.route_info is not None and self.route_info.opposite_exists)
            if busy:
                self._busy_step = 0
                self._set_status(text, STATUS_BUSY)
                self._busy_timer.start()
            else:
                self._busy_timer.stop()

        def _animate_status(self):
            if self._busy_text is not None:
                self._busy_step = (self._busy_step + 1) % 4
                self._set_status(self._busy_text + "." * self._busy_step, STATUS_BUSY)

        @staticmethod
        def _fill(combo, items, select=None):
            """items: [(text, data)]. Selects the entry whose data equals `select`."""
            combo.blockSignals(True)
            combo.clear()
            for text, data in items:
                combo.addItem(text, data)
            index = combo.findData(select) if select is not None else -1
            combo.setCurrentIndex(index)
            if index < 0 and combo.isEditable():
                combo.clearEditText()
            combo.blockSignals(False)
            return index >= 0

        @staticmethod
        def _measure_of(combo):
            text = combo.currentText().strip()
            if not text:
                raise ValueError("Enter a measure.")
            try:
                return float(text.split()[0])
            except ValueError:
                raise ValueError(f"'{text}' is not a measure.") from None

        # -- service chain --------------------------------------------------
        def _connect(self):
            self._run_async("CONNECTING", self.service.jurisdictions, self._connected)

        def _connected(self, names):
            self.log(f"Connected to VDOT LRS Route Master: {len(names)} jurisdictions.")
            items = [(name or STATEWIDE_LABEL, name) for name in names]
            select = DEFAULT_JURISDICTION
            if select is None and self.settings.value("jurisdiction_saved") == "1":
                select = self.settings.value("jurisdiction", "")
            if self._fill(self.jurisdiction, items, select):
                self._load_routes()
            else:
                self.route_hint.setText("Choose a jurisdiction.")

        def _load_routes(self):
            self._clear_route()
            self._fill(self.route, [])
            jurisdiction = self.jurisdiction.currentData()
            if jurisdiction is None:
                return
            self.route_hint.setText("Loading routes\u2026")
            self._run_async("LOADING", lambda: self.service.routes(jurisdiction),
                            lambda result: self._routes_loaded(jurisdiction, result),
                            self.service.cached_routes(jurisdiction))

        def _routes_loaded(self, jurisdiction, result):
            routes, statewide = result
            self.settings.setValue("jurisdiction", jurisdiction)
            self.settings.setValue("jurisdiction_saved", "1")
            if jurisdiction:
                self.log(f"{len(routes)} route(s) for {jurisdiction}: {statewide} statewide "
                         f"(Interstate, US, VA, ramps) crossing it, listed first, and "
                         f"{len(routes) - statewide} of its own.")
            else:
                self.log(f"{len(routes)} statewide route(s), Interstate, US and VA first.")
            if self._fill(self.route, routes, DEFAULT_ROUTE_ID):
                self._load_route()
            else:
                self.route_hint.setText("Choose a route.")

        def _clear_route(self):
            self.route_info = None
            self.geometry_uor = None
            self.route_preview, self.event_preview = [], []
            self._last_event_key = None
            self.route_hint.setText("")
            self._set_event_hint("")
            self.from_location.setText("")
            self.to_location.setText("")
            for combo in (self.from_measure, self.to_measure):
                self._fill(combo, [])
            for button in (self.pick_button, self.zoom_button, self.create_button,
                           self.opposite_button, self.swap_button, self.full_button):
                button.setEnabled(False)

        def _load_route(self, name_id=None):
            if name_id is None:
                name_id = self.route.currentData()
            self._clear_route()
            if name_id is None:
                return
            self.route_hint.setText("Loading route\u2026")
            self._run_async("LOADING", lambda: self.service.route(name_id),
                            self._route_loaded, self.service.cached_route(name_id))

        def _route_loaded(self, info):
            if self.route.currentData() != info.name_id:
                # Reached with "Opposite direction": the route may not be in this list.
                index = self.route.findData(info.name_id)
                if index < 0:
                    self.route.addItem(info.common_name, info.name_id)
                    index = self.route.count() - 1
                self.route.setCurrentIndex(index)
            self.route_info = info
            geometry = info.geometry
            vertices = sum(len(path) for path in geometry.paths)
            opposite = f"  \u00b7  opposite {info.opposite_name}" if info.opposite_exists else ""
            self.route_hint.setText(
                f"{info.name_id}  \u00b7  {geometry.m_max:.2f} mi  \u00b7  "
                f"{len(geometry.paths)} part(s){opposite}")
            self.log(f"Loaded {info.common_name}: {vertices} vertices, measures "
                     f"{fmt_measure(geometry.m_min)} to {fmt_measure(geometry.m_max)} mi.")

            stops = [(fmt_measure(m), m) for m in geometry.stops()]
            self._fill(self.from_measure, stops)
            self._fill(self.to_measure, stops)
            if info.name_id == DEFAULT_ROUTE_ID:
                start, end = DEFAULT_FROM_MEASURE, DEFAULT_TO_MEASURE
            else:
                start, end = geometry.m_min, geometry.paths[0][-1][2]
            self.from_measure.setEditText(fmt_measure(start))
            self.to_measure.setEditText(fmt_measure(end))
            self.next_pick = "from"
            self.create_button.setEnabled(IN_MSTN)
            self.swap_button.setEnabled(True)
            self.full_button.setEnabled(True)
            self.opposite_button.setEnabled(info.opposite_exists)
            self.opposite_button.setToolTip(
                f"Load {info.opposite_name}" if info.opposite_exists
                else "The LRS has no separate route for the opposite direction of this road")
            self._reproject()

        def _switch_opposite(self):
            if self.route_info is None or not self.route_info.opposite_exists:
                raise ValueError("This route has no opposite direction in the LRS.")
            self._load_route(self.route_info.opposite_name)

        # -- event ----------------------------------------------------------
        def _event_type(self):
            return EVENT_LINEAR if self.linear_button.isChecked() else EVENT_POINT

        def current_event(self):
            if self.route_info is None:
                raise ValueError("Choose a route first.")
            geometry = self.route_info.geometry
            event_type = self._event_type()

            def snapped(combo):
                measure = self._measure_of(combo)
                if abs(measure - geometry.m_min) < 0.001:
                    return geometry.m_min
                if abs(measure - geometry.m_max) < 0.001:
                    return geometry.m_max
                return measure

            from_measure = snapped(self.from_measure)
            to_measure = snapped(self.to_measure) if event_type == EVENT_LINEAR else from_measure
            return LrsEvent(self.route_info, event_type, from_measure, to_measure)

        def _measures_edited(self):
            is_linear = self._event_type() == EVENT_LINEAR
            for widget in (self.to_label, self.to_measure, self.to_location,
                           self.swap_button, self.full_button):
                widget.setVisible(is_linear)
            if not is_linear:
                self.next_pick = "from"
            self._sync_pick_state()
            if self.route_info is None:
                return
            # Choosing from the list and then leaving the box both report an edit; the
            # preview is rebuilt only when the event really changed.
            key = (self.route_info.name_id, is_linear, self.from_measure.currentText().strip(),
                   self.to_measure.currentText().strip() if is_linear else "")
            if key == self._last_event_key:
                return
            self._last_event_key = key
            self.offer_bar.hide()
            self.event_preview = []
            self.from_location.setText("")
            self.to_location.setText("")
            try:
                event = self.current_event()
                event.validate()
            except ValueError as ex:
                self._set_event_hint(str(ex), error=True)      # shown here, not logged
                return

            lat, lon = event.lat_lon(event.from_measure)
            self.from_location.setText(f"{lat:.4f}, {lon:.4f}")
            if is_linear:
                lat, lon = event.lat_lon(event.to_measure)
                self.to_location.setText(f"{lat:.4f}, {lon:.4f}")
                parts = len(self.route_info.geometry.slice(event.from_measure, event.to_measure))
                gap = f"  \u00b7  crosses a gap: {parts} elements" if parts > 1 else ""
                self._set_event_hint(
                    f"Event length {event.length:.3f} mi of "
                    f"{self.route_info.geometry.m_max:.2f} mi{gap}")
            else:
                self._set_event_hint("Point event at the From measure.")
            self._build_event_preview(event)

        def _swap(self):
            start, end = self.from_measure.currentText(), self.to_measure.currentText()
            self.from_measure.setEditText(end)
            self.to_measure.setEditText(start)
            self._measures_edited()

        def _full_route(self):
            if self.route_info is None:
                raise ValueError("Choose a route first.")
            geometry = self.route_info.geometry
            self.from_measure.setEditText(fmt_measure(geometry.m_min))
            self.to_measure.setEditText(fmt_measure(geometry.m_max))
            self.next_pick = "from"
            self._measures_edited()

        def _build_event_preview(self, event):
            if self.geometry_uor is None:
                return
            preview = []
            if event.event_type == EVENT_POINT:
                preview.append(point_element(
                    *self.geometry_uor.point_at(event.from_measure),
                    (POINT_EVENT[0], PREVIEW_FROM[1])))
                self.event_preview = preview
                return
            for part in self.geometry_uor.slice(event.from_measure, event.to_measure):
                preview += linestring_elements(part, PREVIEW_EVENT)
            preview.append(point_element(
                *self.geometry_uor.point_at(event.to_measure), PREVIEW_TO))
            preview.append(point_element(
                *self.geometry_uor.point_at(event.from_measure), PREVIEW_FROM))
            self.event_preview = preview

        # -- drawing --------------------------------------------------------
        def _coordinates_chosen(self):
            self._save_choice(self.coordinates, "coordinates")
            self._reproject()

        def _reproject(self):
            self.geometry_uor = None
            self.route_preview, self.event_preview = [], []
            self._last_event_key = None
            for button in (self.pick_button, self.zoom_button):
                button.setEnabled(False)
            if self.route_info is None:
                return
            if IN_MSTN:
                mode, epsg = self.coordinates.currentData()
                geometry = self.route_info.geometry
                origin = geometry.paths[0][0][:2]
                self.geometry_uor = geometry.mapped(make_projector(mode, epsg, origin, self.log))
                for path in self.geometry_uor.paths:
                    self.route_preview += linestring_elements(
                        [(x, y) for x, y, _m in path], PREVIEW_ROUTE)
                for button in (self.pick_button, self.zoom_button):
                    button.setEnabled(True)
            self._measures_edited()
            if self.geometry_uor is not None:
                self._zoom()
                self._start_pick_tool()

        def _zoom(self):
            if self.geometry_uor is not None:
                zoom_to_extent(self.geometry_uor.extent())

        def _pick_tool_active(self):
            return IN_MSTN and LrsPickTool.active is not None

        def _start_pick_tool(self):
            if self.geometry_uor is None:
                raise ValueError("Choose a route first.")
            if not self._pick_tool_active():
                LrsPickTool.install(self)
                self.log("Pick tool started: move over the view to see the route, data point "
                         "sets From then To, reset ends.")
            self._sync_pick_state()

        def _toggle_pick_tool(self):
            try:
                if self._pick_tool_active():
                    LrsPickTool.active._ExitTool()
                else:
                    self._start_pick_tool()
            finally:
                self._sync_pick_state()

        def _sync_pick_state(self):
            """Show whether the pick tool is running and which measure the next pick sets."""
            active = self._pick_tool_active()
            self.pick_button.setChecked(active)
            self.pick_button.setText("Picking in drawing  \u00b7  click to stop" if active
                                     else "Pick From / To in drawing")
            target = self.next_pick if self._event_type() == EVENT_LINEAR else "from"
            for label, name, text in ((self.from_label, "from", "From"),
                                      (self.to_label, "to", "To")):
                is_next = active and name == target
                label.setText(("\u25b6 " if is_next else "") + f"\u25cf  {text}")
                self._set_flag(label, "next", is_next)

        def pick_prompt(self):
            if self._event_type() == EVENT_POINT:
                return "LRS event: pick the point measure on the route"
            return f"LRS event: pick the {'FROM' if self.next_pick == 'from' else 'TO'} measure on the route"

        def dynamic_elements(self, x, y):
            """Elements to draw for the cursor at (x, y), and the measure under it."""
            if self.geometry_uor is None:
                return [], None
            measure, px, py = self.geometry_uor.locate(x, y)
            cursor = point_element(px, py, PREVIEW_CURSOR)
            return self.route_preview + self.event_preview + [cursor], measure

        def on_pick(self, x, y):
            if self.geometry_uor is None:
                return
            measure, _px, _py = self.geometry_uor.locate(x, y)
            finished = self._event_type() == EVENT_POINT or self.next_pick == "to"
            if self._event_type() == EVENT_POINT:
                self.from_measure.setEditText(fmt_measure(measure))
            elif self.next_pick == "from":
                self.from_measure.setEditText(fmt_measure(measure))
                self.next_pick = "to"
            else:
                self.to_measure.setEditText(fmt_measure(measure))
                self.next_pick = "from"
                start, end = self._measure_of(self.from_measure), measure
                if end < start:
                    self.from_measure.setEditText(fmt_measure(end))
                    self.to_measure.setEditText(fmt_measure(start))
            self.log(f"Picked MP {fmt_measure(measure)}.")
            self._guard(self._measures_edited)
            if finished:
                self._offer_create()

        def _offer_create(self):
            """Ask in the dialog itself (not a modal box) whether to draw the picked event."""
            try:
                event = self.current_event()
                event.validate()
            except ValueError:
                return                                  # the hint already explains why
            if self.geometry_uor is None:
                return
            where = (f"MP {fmt_measure(event.from_measure)} to {fmt_measure(event.to_measure)}"
                     if self._event_type() == EVENT_LINEAR
                     else f"MP {fmt_measure(event.from_measure)}")
            self._show_offer(
                f"Draw {event.route_info.route} {where} in the MicroStation drawing?",
                "Yes, draw it", self._create,
                "Not drawn - adjust the measures or pick again, or click Create alignment.")
            self.log(f"Picked {where} - choose Yes, draw it or No below.")

        def _show_offer(self, text, yes_text, action, declined_log):
            """Confirm inside the dialog. A modal box here would open under the busy cursor."""
            self._offer_action = action
            self._offer_declined = declined_log
            self.offer_text.setText(text)
            self.offer_yes.setText(yes_text)
            self.offer_bar.show()
            self.raise_()
            self.activateWindow()
            self.offer_yes.setFocus()

        def _accept_offer(self):
            self.offer_bar.hide()
            action, self._offer_action = getattr(self, "_offer_action", None), None
            if action is not None:
                action()

        def _decline_offer(self):
            self.offer_bar.hide()
            self._offer_action = None
            self.log(getattr(self, "_offer_declined", "") or "Cancelled.")

        def on_tool_ended(self):
            self.log("Pick tool ended - the drawing preview is hidden.")
            self._sync_pick_state()

        def _create(self):
            event = self.current_event()
            event.validate()
            if self.geometry_uor is None:
                raise ValueError("The route has not been converted to model coordinates.")
            ids = create_event_elements(event, self.geometry_uor, self.labels.currentData(),
                                        self.mileposts.currentData(), self.log)
            self.log(f"SUCCESS - {event.route_info.route} "
                     f"MP {fmt_measure(event.from_measure)} to "
                     f"{fmt_measure(event.to_measure)} created ({len(ids)} element(s)).")

        def _delete_created(self):
            refs, counts = find_created_elements()
            if not refs:
                self.log("Nothing to delete - the active model has no elements from this tool.")
                return
            names = {ITEM_TYPE: "event", PART_LABEL: "label", PART_FROM: "start marker",
                     PART_TO: "stop marker", PART_MILEPOST: "milepost sign"}
            summary = ", ".join(
                f"{count} {names.get(kind, kind.lower())}{'' if count == 1 else 's'}"
                for kind, count in sorted(counts.items(), key=lambda kv: list(names).index(
                    kv[0]) if kv[0] in names else len(names)))

            def delete():
                # Search again: the drawing may have changed while the bar was showing.
                current, _counts = find_created_elements()
                deleted = delete_created_elements(current)
                self.log(f"SUCCESS - deleted {deleted} element(s) placed by this tool "
                         f"({summary}).")
                MessageCenter.ShowInfoMessage(
                    f"VDOT LRS: deleted {deleted} element(s).", summary, False)

            self._show_offer(
                f"Delete {len(refs)} element(s) this tool placed ({summary})? "
                f"Undo in MicroStation brings them back.",
                "Yes, delete", delete, "Nothing deleted.")
            self.log(f"Delete {len(refs)} element(s)? Choose Yes, delete or No below.")

        def closeEvent(self, close_event):
            self._result_timer.stop()
            self._busy_timer.stop()
            if IN_MSTN and LrsPickTool.active is not None:
                try:
                    LrsPickTool.active._ExitTool()
                except Exception:
                    print(traceback.format_exc())
            # Give a request that is still running a moment to finish. The worker is a
            # daemon thread that touches nothing in MicroStation, so it is safe to leave.
            if self._worker is not None and self._worker.is_alive():
                self._worker.join(2.0)
            super().closeEvent(close_event)
            QApplication.quit()


def main():
    if not HAVE_QT:
        print("This script needs PyQt5, which ships with MicroStation's Python.")
        return
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)

    # Running the script again while it is open would nest a second dialog, event loop and
    # pick tool inside the first; bring the open dialog forward instead.
    existing = getattr(sys, "_vdot_lrs_event_dialog", None)
    if existing is not None:
        try:
            if existing.isVisible():
                existing.raise_()
                existing.activateWindow()
                print("[LRS] The VDOT LRS dialog is already open.")
                return
        except RuntimeError:
            pass                                   # the old dialog was already destroyed

    dialog = EventDialog()
    sys._vdot_lrs_event_dialog = dialog
    dialog.show()
    dialog.raise_()
    dialog.activateWindow()

    # Qt owns the loop while the dialog is open; this timer lets MicroStation process
    # its own input, which is what drives the pick tool and its dynamics.
    pump = QTimer()
    if IN_MSTN:
        pump.setInterval(10)
        pump.timeout.connect(PyCadInputQueue.PythonMainLoop)
        pump.start()
    try:
        app.exec_()
    finally:
        pump.stop()
        sys._vdot_lrs_event_dialog = None


if __name__ == "__main__":
    main()
