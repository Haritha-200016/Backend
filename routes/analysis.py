import sys
import json
import base64
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage
from math import radians, sin, cos, sqrt, asin, atan2, degrees
from datetime import datetime
import os
import tempfile
import time
import traceback
import xlsxwriter


# ============================================
# MAP GENERATION LIBRARIES
# ============================================
FOLIUM_AVAILABLE = False
SELENIUM_AVAILABLE = False

try:
    import folium
    FOLIUM_AVAILABLE = True
except ImportError:
    print("⚠️ folium not installed. Maps disabled.", file=sys.stderr)

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager
    SELENIUM_AVAILABLE = True
except ImportError:
    print("⚠️ selenium not installed. Image capture disabled.", file=sys.stderr)

# ============================================
# CONFIGURATION
# ============================================

# ============================================
# GPS VALIDATION SETTINGS
# ============================================
MIN_VALID_LAT = 0
MAX_VALID_LAT = 90
MIN_VALID_LON = 60
MAX_VALID_LON = 90
MAX_GPS_JUMP_METERS = 1000  # 1km
MAX_GPS_JUMP_KM = 1.0

# ============================================
# ROUTE GPX FILES - ADD THIS SECTION
# ============================================
GPX_FILE_PATHS = [
    r"C:\Hari\Project\sample_backend\gpx_file\K2_route_dump_to_excavator.gpx",
    r"C:\Hari\Project\sample_backend\gpx_file\W19_route_excavator_to_dump.gpx",
    r"C:\Hari\Project\sample_backend\gpx_file\weight_bridge_6,7_to_h_point_waste_dump.gpx",
    r"C:\Hari\Project\sample_backend\gpx_file\30_Jul_2026_12_33_02_pm.gpx"
]
ROUTE_BUFFER_METERS = 30  # Distance threshold in meters to consider "On Route"
ROUTE_POINTS = []  # Will store all route points after loading GPX files

# ============================================
# EXCAVATOR POLYGON (Hardcoded locations)
# ============================================
EXCAVATOR_POLYGON = {
    "name": "Excavator Area",
    "buffer": 80,
    "coordinates": [
        [20.410580, 81.064820],
        [20.410420, 81.064980],
        [20.410120, 81.065120],
        [20.409760, 81.065210],
        [20.409350, 81.065280],
        [20.409015, 81.065498],
        [20.408450, 81.065300],
        [20.408300, 81.065170],
        [20.408420, 81.064980],
        [20.408760, 81.064860],
        [20.409250, 81.064760],
        [20.409900, 81.064720],
        [20.408582, 81.064937],
        [20.408582, 81.064937],
        [20.407972, 81.064953],
        [20.407354, 81.064937],
        [20.406648, 81.065143],
        [20.406927, 81.065759],
        [20.407461, 81.065844],
        [20.407982, 81.065810],
        [20.408610, 81.065535]
    ]
}

# ============================================
# CIRCLE-BASED DUMP LOCATIONS
# ============================================
DUMP_LOCATIONS = [
    {"id": 2, "lat": 20.4100060, "lon": 81.0692020, "name": "CSP1", "type": "circle", "radius": 100},
    {"id": 3, "lat": 20.4120070, "lon": 81.0610040, "name": "Dump Point 3", "type": "circle", "radius": 100},
    {"id": 1, "lat": 20.395938, "lon": 81.057697, "name": "H-Point dump wast", "type": "circle", "radius": 100},
    {"id": 7, "lat": 20.410206, "lon":  81.067661, "name": "CSP3", "type": "circle", "radius": 100}
]

# ============================================
# POLYGON-BASED DUMP LOCATION
# ============================================
POLYGON_DUMP = {
    "id": 8,
    "name": "CSP2",
    "type": "polygon",
    "buffer": 10,
    "coordinates": [
        [20.410397, 81.069136],
        [20.409659, 81.068908],
        [20.408986, 81.068988],
        [20.408270, 81.069204],
        [20.407851, 81.069499],
        [20.407742, 81.069823],
        [20.407962, 81.070121],
        [20.408435, 81.070284],
        [20.409158, 81.070216],
        [20.409862, 81.069735]
    ]
}

ALL_DUMP_LOCATIONS = DUMP_LOCATIONS + [POLYGON_DUMP]

DEVICE_TO_HAULER = {
    'd3': '137',
    'd7': '43',
    'd10': '133',
    'd9': '134',
    'd12': '135',
}

HAULER_DEVICES = ['d3', 'd10', 'd9', 'd12']
EXCAVATOR_DEVICE = 'd7'
HAULER_SHEETS = ['133', '134', '135', '137']
DUMP_RADIUS_METERS = 100
EXCAVATOR_BUFFER_METERS = 50

MAINTENANCE_LAT = 20.404977
MAINTENANCE_LON = 81.066210
MAINTENANCE_RADIUS = 80

# ============================================
# TRIP COLORS FOR MAPS
# ============================================
TRIP_COLORS = [
    '#FF0000', '#0066CC', '#00CC00', '#FF9900', '#9900CC',
    '#00CCCC', '#FF0066', '#66CC00', '#CC6600', '#3366FF',
    '#FF33CC', '#33CC33', '#FF6633', '#6633CC', '#33CCFF',
    '#FF3366', '#99CC00', '#CC0066', '#0099CC', '#FF9900'
]

# ============================================
# GPX ROUTE LOADING FUNCTION - ADD THIS
# ============================================
def load_gpx_routes(file_paths):
    """Load track points from multiple GPX files and return a list of (lat, lon)."""
    points = []
    try:
        import gpxpy
    except ImportError:
        print("⚠️ gpxpy not installed. GPX route support disabled.", file=sys.stderr)
        return points

    for path in file_paths:
        try:
            with open(path, 'r', encoding='utf-8') as f:
                gpx = gpxpy.parse(f)
            for track in gpx.tracks:
                for segment in track.segments:
                    for pt in segment.points:
                        points.append((pt.latitude, pt.longitude))
            print(f"✅ Loaded route points from {path}", file=sys.stderr)
        except Exception as e:
            print(f"⚠️ Error loading {path}: {e}", file=sys.stderr)
    
    print(f"🛣️ Total route points loaded: {len(points)}", file=sys.stderr)
    return points

# ============================================
# ROUTE CHECKING FUNCTION - ADD THIS
# ============================================
def is_on_route(lat, lon, route_points, buffer_meters=ROUTE_BUFFER_METERS):
    """Check if a GPS point is within buffer distance of any route point."""
    if not route_points:
        return False, float('inf')
    
    min_dist = float('inf')
    for rlat, rlon in route_points:
        d = haversine_distance(lat, lon, rlat, rlon)
        if d < min_dist:
            min_dist = d
            if min_dist < buffer_meters:  # Early exit if already within buffer
                break
    
    return min_dist <= buffer_meters, min_dist

# ============================================
# GPS VALIDATION FUNCTIONS
# ============================================

def is_valid_gps(lat, lon):
    if lat == 0 or lon == 0:
        return False
    if lat < MIN_VALID_LAT or lat > MAX_VALID_LAT:
        return False
    if lon < MIN_VALID_LON or lon > MAX_VALID_LON:
        return False
    return True

# ============================================
# POLYGON HELPER FUNCTIONS - ONLY USED FOR DETECTION
# ============================================

def haversine_distance(lat1, lon1, lat2, lon2):
    """ONLY used for polygon/buffer detection, NOT for trip distances"""
    R = 6371000
    lat1, lon1, lat2, lon2 = map(radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = sin(dlat/2)**2 + cos(lat1) * cos(lat2) * sin(dlon/2)**2
    c = 2 * asin(sqrt(a))
    return R * c

def point_in_polygon(lat, lon, polygon_coords):
    x = lon
    y = lat
    inside = False
    n = len(polygon_coords)
    
    for i in range(n):
        x1 = polygon_coords[i][1]
        y1 = polygon_coords[i][0]
        x2 = polygon_coords[(i + 1) % n][1]
        y2 = polygon_coords[(i + 1) % n][0]
        
        if (y1 == y2 and y == y1 and min(x1, x2) <= x <= max(x1, x2)):
            return True
        
        if ((y1 > y) != (y2 > y)):
            x_intersect = x1 + (y - y1) * (x2 - x1) / (y2 - y1)
            if x <= x_intersect:
                inside = not inside
    
    return inside

def point_on_segment_with_distance(lat, lon, p1, p2, buffer_m):
    lat1, lon1 = p1
    lat2, lon2 = p2
    
    dx = lat2 - lat1
    dy = lon2 - lon1
    
    if dx == 0 and dy == 0:
        dist = haversine_distance(lat, lon, lat1, lon1)
        return dist <= buffer_m, dist
    
    t = ((lat - lat1) * dx + (lon - lon1) * dy) / (dx * dx + dy * dy)
    t = max(0, min(1, t))
    
    closest_lat = lat1 + t * dx
    closest_lon = lon1 + t * dy
    
    dist = haversine_distance(lat, lon, closest_lat, closest_lon)
    return dist <= buffer_m, dist

def point_in_polygon_with_buffer(lat, lon, polygon_coords, buffer_m):
    if point_in_polygon(lat, lon, polygon_coords):
        return True, "inside_polygon", 0
    
    n = len(polygon_coords)
    min_dist = float('inf')
    for i in range(n):
        p1 = polygon_coords[i]
        p2 = polygon_coords[(i + 1) % n]
        is_within, dist = point_on_segment_with_distance(lat, lon, p1, p2, buffer_m)
        if dist < min_dist:
            min_dist = dist
        if is_within:
            return True, f"within_buffer_{dist:.1f}m", dist
    
    return False, f"outside_{min_dist:.1f}m", min_dist

def is_in_polygon_dump(lat, lon):
    if not is_valid_gps(lat, lon):
        return False
    return point_in_polygon_with_buffer(
        lat, lon, 
        POLYGON_DUMP["coordinates"], 
        POLYGON_DUMP["buffer"]
    )[0]

def is_in_excavator_area(lat, lon):
    if not is_valid_gps(lat, lon):
        return False, "invalid", 0
    return point_in_polygon_with_buffer(
        lat, lon,
        EXCAVATOR_POLYGON["coordinates"],
        EXCAVATOR_BUFFER_METERS
    )


def detect_shift_from_data(trips_data):
    """Detect shift (A, B, C) based on the first record timestamp"""
    
    # Find the earliest timestamp from all haulers
    earliest_time = None
    
    for device_id, device_data in trips_data.items():
        first_rec = device_data.get('first_record')
        if first_rec:
            if earliest_time is None or first_rec < earliest_time:
                earliest_time = first_rec
    
    if not earliest_time:
        return "Unknown"
    
    # Parse the time
    try:
        # Try different formats
        for fmt in ['%Y-%m-%d %H:%M:%S', '%Y-%m-%dT%H:%M:%S', '%Y-%m-%d %H:%M:%S.%f']:
            try:
                dt = datetime.strptime(str(earliest_time), fmt)
                break
            except:
                continue
        else:
            return "Unknown"
        
        hour = dt.hour
        
        # Classify shift based on hour
        if 6 <= hour < 14:
            return "A (6 AM - 2 PM)"
        elif 14 <= hour < 22:
            return "B (2 PM - 10 PM)"
        else:
            return "C (10 PM - 6 AM)"
            
    except Exception as e:
        print(f"⚠️ Error detecting shift: {e}", file=sys.stderr)
        return "Unknown"

# ============================================
# HELPER FUNCTIONS
# ============================================

def get_hauler_number(device_id):
    if not device_id:
        return device_id
    return DEVICE_TO_HAULER.get(device_id.lower(), device_id)

def is_hauler(device_id):
    if not device_id:
        return False
    return device_id.lower() in HAULER_DEVICES

def is_excavator(device_id):
    if not device_id:
        return False
    return device_id.lower() == EXCAVATOR_DEVICE

def is_in_maintenance_area(lat, lon):
    if not is_valid_gps(lat, lon):
        return False
    return haversine_distance(lat, lon, MAINTENANCE_LAT, MAINTENANCE_LON) <= MAINTENANCE_RADIUS

def is_at_dump(lat, lon):
    if not is_valid_gps(lat, lon):
        return False, None, 0
    
    for dump in DUMP_LOCATIONS:
        dist = haversine_distance(lat, lon, dump['lat'], dump['lon'])
        if dist <= dump['radius']:
            return True, dump, dist
    
    if is_in_polygon_dump(lat, lon):
        return True, POLYGON_DUMP, 0
    
    return False, None, 0

def get_location_type(lat, lon):
    if not is_valid_gps(lat, lon):
        return 'unknown', f'Invalid GPS ({lat:.4f}, {lon:.4f})'
    
    if is_in_maintenance_area(lat, lon):
        return 'maintenance', 'Maintenance Area'
    
    at_dump, dump, dist = is_at_dump(lat, lon)
    if at_dump:
        return 'dump', dump['name']
    
    in_exc, status, dist = is_in_excavator_area(lat, lon)
    if in_exc:
        return 'excavator', f'Excavator Area ({status})'
    
    return 'unknown', f'Location ({lat:.4f}, {lon:.4f})'

def calculate_duration(start_time, end_time):
    try:
        for fmt in ['%Y-%m-%d %H:%M:%S', '%Y-%m-%dT%H:%M:%S', '%Y-%m-%d %H:%M:%S.%f']:
            try:
                start = datetime.strptime(str(start_time), fmt)
                end = datetime.strptime(str(end_time), fmt)
                diff_seconds = (end - start).total_seconds()
                if diff_seconds >= 0:
                    return round(diff_seconds / 60, 1)
            except:
                continue
        return 0
    except:
        return 0

def calculate_total_distance_for_excavator(df):
    """Calculate actual path distance from consecutive GPS points"""
    if df.empty or len(df) < 2:
        return 0.0
    
    total_distance_km = 0.0
    valid_segments = 0
    max_jump_km = 3.0
    
    # Sort by timestamp
    if 'time' in df.columns:
        df = df.sort_values('time').reset_index(drop=True)
    
    for i in range(1, len(df)):
        lat1 = df.iloc[i-1].get('lat', 0)
        lon1 = df.iloc[i-1].get('lon', 0)
        lat2 = df.iloc[i].get('lat', 0)
        lon2 = df.iloc[i].get('lon', 0)
        
        if not is_valid_gps(lat1, lon1) or not is_valid_gps(lat2, lon2):
            continue
        
        segment_distance_km = haversine_distance(lat1, lon1, lat2, lon2) / 1000
        
        if segment_distance_km > max_jump_km:
            continue
        
        total_distance_km += segment_distance_km
        valid_segments += 1
    
    return round(total_distance_km, 2)

def analyze_excavator_data(records):
    """Analyze excavator data from records"""
    excavator_records = {}
    
    for record in records:
        device_id = record.get('device_id')
        if not device_id:
            continue
        
        if not is_excavator(device_id):
            continue
        
        lat = float(record.get('lat', 0))
        lon = float(record.get('lon', 0))
        timestamp = record.get('time', '')
        vibration = float(record.get('vibration', 0))
        
        if not is_valid_gps(lat, lon):
            continue
        
        if device_id not in excavator_records:
            excavator_records[device_id] = {
                'points': [],
                'first_record': None,
                'last_record': None
            }
        
        excavator_records[device_id]['points'].append({
            'lat': lat,
            'lon': lon,
            'time': timestamp,
            'vibration': vibration
        })
        
        if excavator_records[device_id]['first_record'] is None:
            excavator_records[device_id]['first_record'] = timestamp
        excavator_records[device_id]['last_record'] = timestamp
    
    # Analyze each excavator
    results = {}
    for device_id, data in excavator_records.items():
        df = pd.DataFrame(data['points'])
        
        if df.empty:
            continue
        
        # Calculate metrics
        total_distance = calculate_total_distance_for_excavator(df)
        
        # Idle from vibration
        vibration_values = df['vibration'].fillna(0)
        idle_count = (vibration_values < 0.09).sum()
        total_count = len(df)
        idle_pct = (idle_count / total_count) * 100 if total_count > 0 else 0
        
        # Maintenance percentage
        maintenance_count = 0
        for _, row in df.iterrows():
            if is_in_maintenance_area(row['lat'], row['lon']):
                maintenance_count += 1
        maintenance_pct = (maintenance_count / total_count) * 100 if total_count > 0 else 0
        
        # Runtime
        runtime_pct = max(0, 100 - idle_pct - maintenance_pct)
        
        # Status
        if maintenance_pct > 50:
            status = "In Maintenance Area"
        elif idle_pct > 80:
            status = "Stationary (High Idle)"
        elif idle_pct > 50:
            status = "Stationary (Moderate Idle)"
        elif idle_pct > 30:
            status = "Working (Some Idle)"
        else:
            status = "Active Operation"
        
        results[device_id] = {
            'equipment_id': get_hauler_number(device_id),
            'total_distance': total_distance,
            'idle_percentage': round(idle_pct, 1),
            'runtime_percentage': round(runtime_pct, 1),
            'maintenance_percentage': round(maintenance_pct, 1),
            'operating_status': status,
            'total_points': total_count
        }
    
    return results

def calculate_duration(start_time, end_time):
    """Calculate duration in minutes between two timestamps"""
    try:
        # If already timedelta or datetime objects
        if hasattr(start_time, 'total_seconds') and hasattr(end_time, 'total_seconds'):
            diff_seconds = (end_time - start_time).total_seconds()
            return round(diff_seconds / 60, 1) if diff_seconds > 0 else 0
        
        # Convert to string for parsing
        str_start = str(start_time)
        str_end = str(end_time)
        
        # If empty strings, return 0
        if not str_start or not str_end:
            return 0
        
        # Try different datetime formats
        for fmt in ['%Y-%m-%d %H:%M:%S', '%Y-%m-%dT%H:%M:%S', '%Y-%m-%d %H:%M:%S.%f', '%Y-%m-%d %H:%M:%S.%f%z']:
            try:
                start = datetime.strptime(str_start, fmt)
                end = datetime.strptime(str_end, fmt)
                diff_seconds = (end - start).total_seconds()
                if diff_seconds > 0:
                    return round(diff_seconds / 60, 1)
            except:
                continue
        
        # Try pandas if available
        try:
            import pandas as pd
            start = pd.to_datetime(str_start)
            end = pd.to_datetime(str_end)
            diff_seconds = (end - start).total_seconds()
            if diff_seconds > 0:
                return round(diff_seconds / 60, 1)
        except:
            pass
            
        return 0
    except Exception as e:
        return 0

def analyze_trips(records):
    trips_data = {}
    all_point_data = []
    
    # First pass: collect all records per hauler with timestamps
    hauler_records = {}
    
    for record in records:
        device_id = record.get('device_id')
        if not device_id:
            continue
        if is_excavator(device_id):
            continue
        if not is_hauler(device_id):
            continue
        
        lat = record.get('lat', 0)
        lon = record.get('lon', 0)
        timestamp = record.get('time', '')
        
        try:
            lat = float(lat)
            lon = float(lon)
        except:
            continue
        
        if not is_valid_gps(lat, lon):
            continue
        
        if device_id not in hauler_records:
            hauler_records[device_id] = []
        
        hauler_records[device_id].append(record)
    
    # Process each hauler's records in chronological order
    for device_id, records_list in hauler_records.items():
        # Sort records by timestamp
        records_list.sort(key=lambda x: x.get('time', ''))
        
        # Initialize hauler data
        hauler = {
            'hauler': get_hauler_number(device_id),
            'first_record': None,
            'last_record': None,
            'trips': [],
            'current_trip': None,
            'trip_points': [],
            'all_gps_points': [],
            'maintenance_points': [],
            'at_excavator': False,
            'at_dump': False,
            'excavator_start_time': None,
            'dump_start_time': None,
            'trip_start_time': None,
            'current_location': None,
            'location_start_time': None,
            'location_start_lat': None,
            'location_start_lon': None,
            'total_maintenance_minutes': 0,
            'total_unknown_minutes': 0,
            'total_dump_minutes': 0,
            'total_excavator_minutes': 0,
            'total_loaded_distance_km': 0,
            'total_trip_distance_km': 0,
            'total_maintenance_distance_km': 0,
            'last_gps_point': None,
            'trip_ended_in_maintenance': False,
            'trip_ended_in_unknown': False,
            'trip_visited_maintenance': False,
            'in_trip': False,
            'last_trip_point': None,
            'excavator_start_lat': None,
            'excavator_start_lon': None,
            'trip_unknown_points': [],
            'trip_distance': 0,
            'trip_loaded_distance': 0,
            'trip_start_distance': 0,
            'trip_first_excavator_distance': 0,
            'trip_first_dump_distance': 0,
            'trip_last_dump_distance': 0,
            'trip_end_distance': 0,
            'trip_started_with_buffer': False,
            'trip_dump_location': None,
            'first_dump_found': False,
            'trip_accumulated_distance': 0,
            'last_location_type': None,
            'was_in_maintenance': False,
            'trip_started_in_maintenance': False,
            'trip_started_in_maintenance_distance': 0,
            'pending_maintenance_start': False,
            'maintenance_to_excavator_distance': 0,
            'maint_to_ex_saved': False,
            'last_maintenance_timestamp': None,
            'last_maintenance_lat': None,
            'last_maintenance_lon': None,
            'last_maintenance_distance': 0,
            'maint_to_ex_distance_calculated': 0,
            'was_in_maintenance_prev': False,
            'maintenance_entry_time': None,
            'maintenance_entry_lat': None,
            'maintenance_entry_lon': None,
            'maintenance_periods': [],
            'prev_timestamp': None,
            'prev_location_type': None,
            'prev_lat': None,
            'prev_lon': None,
            'prev_vibration': 0,
            'idle_time': 0,
            'productive_idle_minutes': 0,
            'productive_idle_engine_on_minutes': 0,
            'productive_idle_engine_off_minutes': 0,
            'nonproductive_onroute_idle_minutes': 0,
            'nonproductive_onroute_engine_on_minutes': 0,
            'nonproductive_onroute_engine_off_minutes': 0,
            'nonproductive_offroute_idle_minutes': 0,
            'nonproductive_offroute_engine_on_minutes': 0,
            'nonproductive_offroute_engine_off_minutes': 0,
            'running_hours_minutes': 0,
            'travel_time_minutes': 0,
            'excavator_wait_minutes': 0,
            'dump_wait_minutes': 0,
            'prev_is_moving': False,
            'prev_on_route': False,  # ADDED: Track previous point's route status
            'movement_threshold_meters': 10,
            'total_travel_distance_km': 0,
            'total_roaming_distance_km': 0,
            'current_trip_max_altitude': 0,
            'lift_base_altitude': 390,
            'trip_start_fuel_raw': None,
            'last_fuel_level': 0,
            'avg_lift': 0,
            'trip_dump_points': [],
            'excavator_return_point': None,
            'offroute_time': 0.0,
            'current_trip_points': [],
        }
        
        trips_data[device_id] = hauler
        
        # Process each record
        for i, record in enumerate(records_list):
            lat = float(record.get('lat', 0))
            lon = float(record.get('lon', 0))
            timestamp = str(record.get('time', ''))
            
            if hauler['prev_lat'] is not None and hauler['prev_lon'] is not None:
                distance = haversine_distance(hauler['prev_lat'], hauler['prev_lon'], lat, lon) / 1000
            else:
                distance = 0
            
            fuel_consumption = record.get('fuel_consumption', 0)
            try:
                fuel_consumption = float(fuel_consumption) if fuel_consumption else 0
            except:
                fuel_consumption = 0
                
            altitude = record.get('altitude', record.get('alt', 0))
            try:
                altitude = float(altitude) if altitude else 0
            except:
                altitude = 0    
                
            pitch = record.get('pitch', 0)
            roll = record.get('roll', 0)
            vibration = record.get('vibration', 0)
            fuel = record.get('fuel', 0)
            speed = record.get('speed', 0)
            
            try:
                pitch = float(pitch) if pitch else 0
                roll = float(roll) if roll else 0
                vibration = float(vibration) if vibration else 0
                fuel = float(fuel) if fuel else 0
                speed = float(speed) if speed else 0
            except:
                continue
            
            if not is_valid_gps(lat, lon):
                continue
            
            loc_type, loc_name = get_location_type(lat, lon)
            
            is_inside_polygon = 'inside_polygon' in loc_name.lower()
            is_within_buffer = 'within_buffer' in loc_name.lower()
            
            # ===== CHECK ROUTE STATUS =====
            on_route = False
            dist_to_route = float('inf')
            route_status = "Off Route"
            
            if ROUTE_POINTS:
                on_route, dist_to_route = is_on_route(lat, lon, ROUTE_POINTS, ROUTE_BUFFER_METERS)
                route_status = "On Route" if on_route else "Off Route"
            # ===== END ROUTE CHECK =====
            
            # ===== DETECT MOVEMENT =====
            is_moving = False
            movement_distance = 0
            if hauler['prev_lat'] is not None and hauler['prev_lon'] is not None:
                movement_distance = haversine_distance(hauler['prev_lat'], hauler['prev_lon'], lat, lon)
                is_moving = movement_distance > hauler['movement_threshold_meters']
            
            # ===== DETECT ENGINE ON/OFF =====
            is_engine_on = vibration >= 0.09 if vibration is not None else True
            
            point_info = {
                'hauler': get_hauler_number(device_id),
                'timestamp': timestamp,
                'lat': lat,
                'lon': lon,
                'location_type': loc_type,
                'location_name': loc_name,
                'distance': distance,
                'fuel_consumption': fuel_consumption,
                'pitch': pitch,
                'roll': roll,
                'vibration': vibration,
                'fuel': fuel,
                'speed': speed,
                'is_inside_polygon': is_inside_polygon,
                'is_within_buffer': is_within_buffer,
                'altitude': altitude,
                'is_moving': is_moving,
                'movement_distance_m': movement_distance,
                'is_engine_on': is_engine_on,
                'on_route': on_route,
                'dist_to_route_m': round(dist_to_route, 1) if dist_to_route != float('inf') else None,
                'route_status': route_status
            }
            all_point_data.append(point_info)
            
            if hauler['first_record'] is None:
                hauler['first_record'] = timestamp
            hauler['last_record'] = timestamp
            
            hauler['all_gps_points'].append({
                'lat': lat,
                'lon': lon,
                'time': timestamp,
                'type': loc_type,
                'name': loc_name,
                'distance': distance,
                'is_inside_polygon': is_inside_polygon,
                'is_within_buffer': is_within_buffer,
                'is_moving': is_moving,
                'movement_distance_m': movement_distance,
                'altitude': altitude,
                'vibration': vibration,
                'is_engine_on': is_engine_on,
                'on_route': on_route,
                'dist_to_route_m': round(dist_to_route, 1) if dist_to_route != float('inf') else None,
                'route_status': route_status
            })
            
            # ============================================
            # MAINTENANCE TIME TRACKING
            # ============================================
            if loc_type == 'maintenance':
                if not hauler['was_in_maintenance_prev']:
                    hauler['maintenance_entry_time'] = timestamp
                    hauler['maintenance_entry_lat'] = lat
                    hauler['maintenance_entry_lon'] = lon
                    hauler['was_in_maintenance_prev'] = True
                    print(f"🔧 {device_id}: MAINTENANCE ENTRY at {timestamp}", file=sys.stderr)
            else:
                if hauler['was_in_maintenance_prev'] and hauler['maintenance_entry_time'] is not None:
                    duration = calculate_duration(hauler['maintenance_entry_time'], timestamp)
                    if duration > 0:
                        hauler['total_maintenance_minutes'] += duration
                        hauler['maintenance_periods'].append({
                            'entry': hauler['maintenance_entry_time'],
                            'exit': timestamp,
                            'duration_minutes': duration
                        })
                        print(f"🔧 {device_id}: MAINTENANCE EXIT at {timestamp} - Duration: {duration:.1f} min", file=sys.stderr)
                    
                    hauler['maintenance_entry_time'] = None
                    hauler['maintenance_entry_lat'] = None
                    hauler['maintenance_entry_lon'] = None
                    hauler['was_in_maintenance_prev'] = False
            
            # ============================================
            # IDLE CATEGORIES (uses previous values)
            # ============================================
            if hauler['prev_timestamp'] is not None:
                time_diff = calculate_duration(hauler['prev_timestamp'], timestamp)
                if time_diff > 0 and time_diff < 1.0:
                    prev_type = hauler['prev_location_type']
                    prev_is_moving = hauler.get('prev_is_moving', False)
                    prev_is_engine_on = hauler.get('prev_is_engine_on', True)
                    
                    if prev_type == 'excavator':
                        if not prev_is_moving:
                            hauler['productive_idle_minutes'] += time_diff
                            hauler['excavator_wait_minutes'] += time_diff
                            hauler['running_hours_minutes'] += time_diff
                            if prev_is_engine_on:
                                hauler['productive_idle_engine_on_minutes'] += time_diff
                            else:
                                hauler['productive_idle_engine_off_minutes'] += time_diff
                        
                    elif prev_type == 'dump':
                        if not prev_is_moving:
                            hauler['productive_idle_minutes'] += time_diff
                            hauler['dump_wait_minutes'] += time_diff
                            hauler['running_hours_minutes'] += time_diff
                            if prev_is_engine_on:
                                hauler['productive_idle_engine_on_minutes'] += time_diff
                            else:
                                hauler['productive_idle_engine_off_minutes'] += time_diff
                    
                    elif prev_type == 'unknown':
                        if prev_is_moving:
                            hauler['travel_time_minutes'] += time_diff
                            hauler['running_hours_minutes'] += time_diff
                            hauler['total_travel_distance_km'] += movement_distance / 1000
                        else:
                            # ===== FIXED: Use previous point's route status =====
                            if hauler.get('prev_on_route', False):
                                hauler['nonproductive_onroute_idle_minutes'] += time_diff
                                hauler['running_hours_minutes'] += time_diff
                                if prev_is_engine_on:
                                    hauler['nonproductive_onroute_engine_on_minutes'] += time_diff
                                else:
                                    hauler['nonproductive_onroute_engine_off_minutes'] += time_diff
                            else:
                                hauler['nonproductive_offroute_idle_minutes'] += time_diff
                                if prev_is_engine_on:
                                    hauler['nonproductive_offroute_engine_on_minutes'] += time_diff
                                else:
                                    hauler['nonproductive_offroute_engine_off_minutes'] += time_diff
                    
                    elif prev_type == 'maintenance':
                        pass
            
            # ============================================
            # TRACK LAST MAINTENANCE POINT
            # ============================================
            if loc_type == 'maintenance':
                hauler['last_maintenance_timestamp'] = timestamp
                hauler['last_maintenance_lat'] = lat
                hauler['last_maintenance_lon'] = lon
                hauler['last_maintenance_distance'] = distance
                hauler['was_in_maintenance'] = True
                hauler['maint_to_ex_saved'] = False
            
            # ===== CALCULATE M→Ex DISTANCE =====
            if loc_type == 'excavator' and hauler.get('was_in_maintenance', False) and hauler.get('current_trip') is None:
                if not hauler.get('maint_to_ex_saved', False):
                    maint_to_ex_dist = 0
                    last_maint_time = hauler.get('last_maintenance_timestamp')
                    
                    if last_maint_time:
                        for point in hauler['all_gps_points']:
                            if point.get('time') and point.get('time') > last_maint_time:
                                hop_dist = point.get('distance', 0)
                                if hop_dist > 0 and hop_dist <= 1:
                                    maint_to_ex_dist += hop_dist
                    
                    if maint_to_ex_dist > 0:
                        hauler['maintenance_to_excavator_distance'] = maint_to_ex_dist
                        hauler['maint_to_ex_distance_calculated'] = maint_to_ex_dist
                        hauler['pending_maintenance_start'] = True
                        hauler['maint_to_ex_saved'] = True
                        print(f"🚚 {device_id}: M→Ex distance calculated: {maint_to_ex_dist:.2f} km", file=sys.stderr)
            
            # ============================================
            # ORIGINAL LOCATION CHANGE TRACKING
            # ============================================
            if hauler['current_location'] != loc_name:
                hauler['current_location'] = loc_name
                hauler['location_start_time'] = timestamp
                hauler['location_start_lat'] = lat
                hauler['location_start_lon'] = loc_name
            
            # ============================================
            # TRIP LOGIC
            # ============================================
            if loc_type == 'maintenance' and hauler['current_trip'] is not None:
                hauler['trip_visited_maintenance'] = True
                hauler['maintenance_points'].append((lat, lon, timestamp, distance))
            
            if loc_type == 'maintenance' and hauler['at_dump'] and hauler['current_trip'] is not None:
                hauler['trip_ended_in_maintenance'] = True
            
            # ============================================
            # START TRIP - At Excavator
            # ============================================
            if loc_type == 'excavator' and not hauler['at_excavator'] and hauler['current_trip'] is None:
                if is_inside_polygon or is_within_buffer:
                    if hauler.get('pending_maintenance_start', False):
                        hauler['trip_started_in_maintenance'] = True
                        hauler['trip_started_in_maintenance_distance'] = hauler.get('maintenance_to_excavator_distance', 0)
                        print(f"🚚 {device_id}: TRIP STARTED FROM MAINTENANCE! M→Ex: {hauler['trip_started_in_maintenance_distance']:.2f} km", file=sys.stderr)
                    else:
                        hauler['trip_started_in_maintenance'] = False
                        hauler['trip_started_in_maintenance_distance'] = 0
                    
                    hauler['at_excavator'] = True
                    hauler['in_trip'] = True
                    hauler['trip_start_time'] = timestamp
                    hauler['excavator_start_time'] = timestamp
                    hauler['trip_points'] = []
                    hauler['maintenance_points'] = []
                    hauler['trip_ended_in_maintenance'] = False
                    hauler['trip_ended_in_unknown'] = False
                    hauler['trip_visited_maintenance'] = False
                    hauler['last_trip_point'] = None
                    hauler['trip_unknown_points'] = []
                    hauler['first_dump_found'] = False
                    hauler['current_trip_points'] = []
                    
                    hauler['trip_dump_points'] = []
                    hauler['excavator_return_point'] = None
                    
                    hauler['trip_start_fuel_raw'] = None
                    hauler['last_fuel_level'] = 0
                    
                    hauler['was_in_maintenance'] = False
                    hauler['pending_maintenance_start'] = False
                    
                    hauler['trip_first_excavator_distance'] = 0
                    hauler['trip_start_distance'] = 0
                    hauler['trip_started_with_buffer'] = not is_inside_polygon
                    
                    hauler['current_trip_max_altitude'] = 0
                    
                    hauler['current_trip'] = {
                        'excavator_arrival': timestamp,
                        'excavator_lat': lat,
                        'excavator_lon': lon,
                        'started_with_buffer': not is_inside_polygon,
                        'used_inside_polygon': is_inside_polygon,
                        'dump_points': [],
                        'excavator_return_lat': None,
                        'excavator_return_lon': None,
                        'excavator_return_time': None,
                        'offroute_time': 0.0,
                    }
                    
                    # Reset offroute_time for new trip
                    hauler['offroute_time'] = 0.0
                    
                    # ===== STORE INITIAL WAIT VALUES FOR THIS TRIP =====
                    hauler['trip_start_excavator_wait'] = hauler.get('excavator_wait_minutes', 0)
                    hauler['trip_start_dump_wait'] = hauler.get('dump_wait_minutes', 0)
                    hauler['trip_start_productive_idle'] = hauler.get('productive_idle_minutes', 0)
                    
                    hauler['trip_points'].append((lat, lon, timestamp, distance, is_inside_polygon, is_within_buffer))
                    hauler['last_trip_point'] = (lat, lon, timestamp, distance, is_inside_polygon, is_within_buffer)
                    
                    hauler['trip_accumulated_distance'] = 0
                    
                    if is_inside_polygon:
                        print(f"🚚 {device_id}: TRIP START - Inside Excavator Polygon at {timestamp}", file=sys.stderr)
                    else:
                        print(f"🚚 {device_id}: TRIP START - Excavator Buffer (50m) at {timestamp}", file=sys.stderr)
            
            # ============================================
            # TRACK ALL POINTS DURING TRIP
            # ============================================
            elif hauler['current_trip'] is not None:
                hauler['trip_points'].append((lat, lon, timestamp, distance, is_inside_polygon, is_within_buffer))
                hauler['current_trip_points'].append({  # ← ADD THIS BLOCK (missing!)
                    'lat': lat,
                    'lon': lon,
                    'time': timestamp,
                    'type': loc_type,
                    'name': loc_name,
                    'distance': distance
                })
                
                if fuel > 0:
                    if hauler.get('trip_start_fuel_raw') is None:
                        hauler['trip_start_fuel_raw'] = fuel
                        print(f"⛽ {device_id}: Trip start fuel raw: {fuel}", file=sys.stderr)
                    hauler['last_fuel_level'] = fuel
                
                if loc_type != 'maintenance' and altitude > 0:
                    if altitude > hauler['current_trip_max_altitude']:
                        hauler['current_trip_max_altitude'] = altitude
                        print(f"📍 {device_id}: New max altitude: {altitude:.1f}m at {loc_type} for trip {len(hauler['trips'])+1}", file=sys.stderr)
                
                if loc_type == 'unknown':
                    hauler['trip_unknown_points'].append((lat, lon, timestamp, distance))
                
                # ===== TRACK OFF-ROUTE TIME =====
                # ONLY count unknown locations that are Off Route
                # Maintenance, dump, and excavator are NOT counted as off-route
                if loc_type == 'unknown' and not on_route:
                    if hauler['prev_timestamp'] is not None:
                        # Calculate time difference using proper duration function
                        time_diff = calculate_duration(hauler['prev_timestamp'], timestamp)
                        if time_diff > 0 and time_diff < 1.0:
                            hauler['current_trip']['offroute_time'] = hauler['current_trip'].get('offroute_time', 0) + time_diff
                            hauler['offroute_time'] = hauler['current_trip']['offroute_time']
                            print(f"🛣️ {device_id}: Off-route at {timestamp} - +{time_diff:.1f}min (total: {hauler['current_trip']['offroute_time']:.1f}min)", file=sys.stderr)
                        else:
                            # Debug: print if time_diff is 0 or too large
                            if time_diff <= 0:
                                print(f"⚠️ {device_id}: time_diff = {time_diff} at {timestamp} (prev: {hauler['prev_timestamp']})", file=sys.stderr)
                            elif time_diff >= 1.0:
                                print(f"⚠️ {device_id}: time_diff = {time_diff} (too large) at {timestamp}", file=sys.stderr)
                    else:
                        # First point of trip - initialize offroute_time
                        if 'offroute_time' not in hauler['current_trip']:
                            hauler['current_trip']['offroute_time'] = 0.0
                
                if distance > 0 and distance <= 1:
                    hauler['trip_accumulated_distance'] += distance
                elif distance > 1:
                    print(f"⚠️ Outlier distance ignored: {distance:.2f}km at {timestamp}", file=sys.stderr)
                
                hauler['last_trip_point'] = (lat, lon, timestamp, distance, is_inside_polygon, is_within_buffer)
                
                # ============================================
                # TRACK WAIT TIMES AT EXCAVATOR AND DUMP
                # ============================================
                if loc_type == 'excavator':
                    if hauler['prev_timestamp'] is not None and hauler['prev_location_type'] == 'excavator':
                        time_diff = calculate_duration(hauler['prev_timestamp'], timestamp)
                        if time_diff > 0 and time_diff < 1.0:
                            hauler['current_trip']['ex_wait_time'] = hauler['current_trip'].get('ex_wait_time', 0) + time_diff
                
                if loc_type == 'dump':
                    if hauler['prev_timestamp'] is not None and hauler['prev_location_type'] == 'dump':
                        time_diff = calculate_duration(hauler['prev_timestamp'], timestamp)
                        if time_diff > 0 and time_diff < 1.0:
                            hauler['current_trip']['dump_wait_time'] = hauler['current_trip'].get('dump_wait_time', 0) + time_diff
            
            # ============================================
            # AT DUMP - Collect ALL dump points
            # ============================================
            if loc_type == 'dump' and hauler['current_trip'] is not None:
                if not hauler['at_dump']:
                    hauler['at_dump'] = True
                    hauler['dump_start_time'] = timestamp
                
                dump_point = {
                    'lat': lat,
                    'lon': lon,
                    'timestamp': timestamp,
                    'distance': hauler['trip_accumulated_distance'],
                    'lead_distance': hauler['trip_accumulated_distance']
                }
                hauler['trip_dump_points'].append(dump_point)
                hauler['current_trip']['dump_points'].append(dump_point)
                
                if not hauler['first_dump_found']:
                    hauler['trip_dump_location'] = loc_name
                    hauler['first_dump_found'] = True
                
                hauler['trip_last_dump_distance'] = hauler['trip_accumulated_distance']
                
                print(f"📍 {device_id}: Dump point {len(hauler['trip_dump_points'])} at {timestamp}, Dist: {hauler['trip_accumulated_distance']:.2f} km", file=sys.stderr)
            
            # ============================================
            # END TRIP - Returned to Excavator (FIRST POINT)
            # ============================================
            elif loc_type == 'excavator' and hauler['at_dump'] and hauler['current_trip'] is not None:
                if is_inside_polygon or is_within_buffer:
                    hauler['excavator_return_point'] = {
                        'lat': lat,
                        'lon': lon,
                        'timestamp': timestamp,
                        'distance': hauler['trip_accumulated_distance']
                    }
                    hauler['current_trip']['excavator_return_lat'] = lat
                    hauler['current_trip']['excavator_return_lon'] = lon
                    hauler['current_trip']['excavator_return_time'] = timestamp
                    
                    hauler['at_dump'] = False
                    hauler['at_excavator'] = False
                    hauler['in_trip'] = False
                    
                    dump_points = hauler['trip_dump_points']
                    
                    if dump_points:
                        dump_points.sort(key=lambda x: x.get('distance', 0))
                        
                        if len(dump_points) >= 5:
                            middle_index = len(dump_points) // 2
                            selected_dump = dump_points[middle_index]
                            print(f"📍 {device_id}: Using MIDDLE dump point ({len(dump_points)} points, index {middle_index})", file=sys.stderr)
                        elif len(dump_points) == 2:
                            selected_dump = dump_points[1]
                            print(f"📍 {device_id}: Using 2nd dump point ({len(dump_points)} points)", file=sys.stderr)
                        elif len(dump_points) == 3:
                            selected_dump = dump_points[1]
                            print(f"📍 {device_id}: Using middle dump point ({len(dump_points)} points)", file=sys.stderr)
                        else:
                            selected_dump = dump_points[0]
                            print(f"📍 {device_id}: Using first dump point ({len(dump_points)} points)", file=sys.stderr)
                        
                        lead_distance_km = selected_dump.get('lead_distance', hauler['trip_first_dump_distance'])
                        dump_lat = selected_dump.get('lat', 0)
                        dump_lon = selected_dump.get('lon', 0)
                        dump_time = selected_dump.get('timestamp', '')
                        
                        print(f"📍 {device_id}: Lead distance based on selected dump point: {lead_distance_km:.2f} km", file=sys.stderr)
                    else:
                        lead_distance_km = hauler['trip_first_dump_distance']
                        dump_lat = hauler['current_trip'].get('dump_lat', 0)
                        dump_lon = hauler['current_trip'].get('dump_lon', 0)
                        dump_time = hauler['current_trip'].get('dump_arrival', '')
                    
                    cycle_distance_km = hauler['trip_accumulated_distance']
                    
                    # Calculate Lead times
                    lead1_time = calculate_duration(hauler['current_trip']['excavator_arrival'], dump_time)
                    lead2_time = calculate_duration(dump_time, timestamp)
                    total_trip_time = calculate_duration(hauler['current_trip']['excavator_arrival'], timestamp)
                    
                    # ===== CALCULATE WAIT TIMES FOR THIS TRIP =====
                    ex_wait = hauler.get('excavator_wait_minutes', 0) - hauler.get('trip_start_excavator_wait', 0)
                    dump_wait = hauler.get('dump_wait_minutes', 0) - hauler.get('trip_start_dump_wait', 0)
                    
                    # ===== GET OFF-ROUTE TIME =====
                    offroute_time = hauler['current_trip'].get('offroute_time', 0)
                    # Path Deviation: YES if offroute time > 5 minutes, else NO
                    path_deviation = 'YES' if offroute_time > 5 else 'NO'
                    
                    print(f"⏱️ {device_id}: TIMES - Lead1: {lead1_time:.1f}min, Lead2: {lead2_time:.1f}min, Total: {total_trip_time:.1f}min", file=sys.stderr)
                    print(f"⏱️ {device_id}: WAIT - Ex: {ex_wait:.1f}min, Dump: {dump_wait:.1f}min", file=sys.stderr)
                    print(f"🛣️ {device_id}: OFF-ROUTE - {offroute_time:.1f}min, Deviation: {path_deviation}", file=sys.stderr)
                    
                    maintenance_distance_km = 0
                    if hauler['trip_visited_maintenance']:
                        maintenance_distance_km = hauler['trip_accumulated_distance'] - hauler['trip_last_dump_distance']
                    
                    start_fuel_raw = hauler.get('trip_start_fuel_raw')
                    end_fuel_raw = hauler.get('last_fuel_level')
                    if start_fuel_raw is None:
                        start_fuel_raw = 0
                    if end_fuel_raw is None:
                        end_fuel_raw = 0
                    
                    if start_fuel_raw > 0 and end_fuel_raw > 0 and start_fuel_raw > end_fuel_raw:
                        fuel_consumption = (start_fuel_raw - end_fuel_raw) / 5
                        print(f"⛽ {device_id}: Fuel - Start: {start_fuel_raw}, End: {end_fuel_raw}, Diff: {start_fuel_raw - end_fuel_raw}, Fuel: {fuel_consumption:.2f}L", file=sys.stderr)
                    else:
                        fuel_consumption = 0
                    
                    started_in_maint = hauler.get('trip_started_in_maintenance', False)
                    started_in_maint_km = hauler.get('trip_started_in_maintenance_distance', 0)
                    
                    max_alt = hauler.get('current_trip_max_altitude', 0)
                    base_alt = hauler.get('lift_base_altitude', 390)
                    lift = max_alt - base_alt if max_alt > base_alt else 0
                    print(f"📊 {device_id}: LIFT CALC - Max Alt: {max_alt:.1f}m, Base: {base_alt}, Lift: {lift:.2f}m", file=sys.stderr)
                    
                    trip = {
                        'excavator_arrival': hauler['current_trip']['excavator_arrival'],
                        'dump_arrival': dump_time,
                        'trip_end': timestamp,
                        'dump_location': hauler['trip_dump_location'],
                        'loaded_distance_km': round(lead_distance_km, 2),
                        'trip_distance_km': round(cycle_distance_km, 2),
                        'ended_in_maintenance': False,
                        'ended_in_unknown': False,
                        'maintenance_distance_km': round(maintenance_distance_km, 2),
                        'visited_maintenance': hauler['trip_visited_maintenance'],
                        'started_with_buffer': hauler['trip_started_with_buffer'],
                        'ended_with_buffer': not is_inside_polygon,
                        'fuel': round(fuel_consumption, 2),
                        'lift': round(lift, 2),
                        'avg_tonnes': 0,
                        'started_in_maintenance': started_in_maint,
                        'started_in_maintenance_km': round(started_in_maint_km, 2),
                        'dump_points_count': len(dump_points),
                        'dump_points': dump_points,
                        'time_at_excavator': round(lead1_time, 1),
                        'wait_at_excavator': round(ex_wait, 1),
                        'time_at_dump': round(lead2_time, 1),
                        'wait_at_dump': round(dump_wait, 1),
                        'total_trip_time': round(total_trip_time, 1),
                        'path_deviation': path_deviation,
                        'path_duration': round(offroute_time, 1),
                        'points': hauler['current_trip_points'].copy()
                    }
                    
                    hauler['total_loaded_distance_km'] += lead_distance_km
                    hauler['total_trip_distance_km'] += cycle_distance_km
                    hauler['trips'].append(trip)
                    
                    print(f"🚚 {device_id}: TRIP END - Lead: {lead_distance_km:.2f}km, Cycle: {cycle_distance_km:.2f}km", file=sys.stderr)
                    print(f"   Lead1: {lead1_time:.1f}min, Lead2: {lead2_time:.1f}min, Total: {total_trip_time:.1f}min", file=sys.stderr)
                    print(f"   Wait Ex: {ex_wait:.1f}min, Wait Dump: {dump_wait:.1f}min", file=sys.stderr)
                    print(f"   Path Deviation: {path_deviation}, Path Duration: {offroute_time:.1f}min", file=sys.stderr)
                    
                    hauler['current_trip'] = None
                    hauler['dump_start_time'] = None
                    hauler['trip_points'] = []
                    hauler['maintenance_points'] = []
                    hauler['last_trip_point'] = None
                    hauler['trip_accumulated_distance'] = 0
                    hauler['excavator_start_lat'] = None
                    hauler['excavator_start_lon'] = None
                    hauler['trip_unknown_points'] = []
                    hauler['trip_started_with_buffer'] = False
                    hauler['trip_first_excavator_distance'] = 0
                    hauler['trip_first_dump_distance'] = 0
                    hauler['trip_last_dump_distance'] = 0
                    hauler['trip_end_distance'] = 0
                    hauler['first_dump_found'] = False
                    hauler['maintenance_to_excavator_distance'] = 0
                    hauler['current_trip_max_altitude'] = 0
                    hauler['trip_start_fuel_raw'] = None
                    hauler['last_fuel_level'] = 0
                    hauler['trip_dump_points'] = []
                    hauler['excavator_return_point'] = None
                    hauler['offroute_time'] = 0.0
            
            # ============================================
            # END TRIP - Maintenance interruption
            # ============================================
            elif loc_type == 'maintenance' and hauler['at_dump'] and hauler['current_trip'] is not None:
                hauler['at_dump'] = False
                hauler['at_excavator'] = False
                hauler['in_trip'] = False
                hauler['trip_ended_in_maintenance'] = True
                
                # ===== ADD THIS LINE =====
                trip_points = hauler.get('current_trip_points', [])
                # ===== END =====
                
                dump_points = hauler['trip_dump_points']
                
                if dump_points:
                    dump_points.sort(key=lambda x: x.get('distance', 0))
                    
                    if len(dump_points) >= 5:
                        middle_index = len(dump_points) // 2
                        selected_dump = dump_points[middle_index]
                    elif len(dump_points) == 2:
                        selected_dump = dump_points[1]
                    elif len(dump_points) == 3:
                        selected_dump = dump_points[1]
                    else:
                        selected_dump = dump_points[0]
                    
                    lead_distance_km = selected_dump.get('lead_distance', hauler['trip_first_dump_distance'])
                    dump_time = selected_dump.get('timestamp', '')
                    print(f"📍 {device_id}: Using selected dump point for maintenance end ({len(dump_points)} points)", file=sys.stderr)
                else:
                    lead_distance_km = hauler['trip_first_dump_distance']
                    dump_time = hauler['current_trip'].get('dump_arrival', '')
                
                cycle_distance_km = hauler['trip_accumulated_distance']
                maintenance_distance_km = hauler['trip_accumulated_distance'] - hauler['trip_last_dump_distance']
                
                lead1_time = calculate_duration(hauler['current_trip']['excavator_arrival'], dump_time)
                lead2_time = calculate_duration(dump_time, timestamp)
                total_trip_time = calculate_duration(hauler['current_trip']['excavator_arrival'], timestamp)
                
                ex_wait = hauler.get('excavator_wait_minutes', 0) - hauler.get('trip_start_excavator_wait', 0)
                dump_wait = hauler.get('dump_wait_minutes', 0) - hauler.get('trip_start_dump_wait', 0)
                
                offroute_time = hauler['current_trip'].get('offroute_time', 0)
                path_deviation = 'YES' if offroute_time > 5 else 'NO'
                
                print(f"⏱️ {device_id}: TIMES - Lead1: {lead1_time:.1f}min, Lead2: {lead2_time:.1f}min, Total: {total_trip_time:.1f}min", file=sys.stderr)
                print(f"⏱️ {device_id}: WAIT - Ex: {ex_wait:.1f}min, Dump: {dump_wait:.1f}min", file=sys.stderr)
                print(f"🛣️ {device_id}: OFF-ROUTE - {offroute_time:.1f}min, Deviation: {path_deviation}", file=sys.stderr)
                
                start_fuel_raw = hauler.get('trip_start_fuel_raw')
                end_fuel_raw = hauler.get('last_fuel_level')
                if start_fuel_raw is None:
                    start_fuel_raw = 0
                if end_fuel_raw is None:
                    end_fuel_raw = 0
                
                if start_fuel_raw > 0 and end_fuel_raw > 0 and start_fuel_raw > end_fuel_raw:
                    fuel_consumption = (start_fuel_raw - end_fuel_raw) / 5
                    print(f"⛽ {device_id}: Fuel - Start: {start_fuel_raw}, End: {end_fuel_raw}, Diff: {start_fuel_raw - end_fuel_raw}, Fuel: {fuel_consumption:.2f}L", file=sys.stderr)
                else:
                    fuel_consumption = 0
                
                started_in_maint = hauler.get('trip_started_in_maintenance', False)
                started_in_maint_km = hauler.get('trip_started_in_maintenance_distance', 0)
                
                max_alt = hauler.get('current_trip_max_altitude', 0)
                base_alt = hauler.get('lift_base_altitude', 390)
                lift = max_alt - base_alt if max_alt > base_alt else 0
                print(f"📊 {device_id}: LIFT CALC - Max Alt: {max_alt:.1f}m, Base: {base_alt}, Lift: {lift:.2f}m", file=sys.stderr)
                
                trip = {
                    'excavator_arrival': hauler['current_trip']['excavator_arrival'],
                    'dump_arrival': dump_time,
                    'trip_end': timestamp,
                    'dump_location': hauler['trip_dump_location'],
                    'loaded_distance_km': round(lead_distance_km, 2),
                    'trip_distance_km': round(cycle_distance_km, 2),
                    'ended_in_maintenance': True,
                    'ended_in_unknown': False,
                    'maintenance_distance_km': round(maintenance_distance_km, 2),
                    'visited_maintenance': hauler['trip_visited_maintenance'],
                    'started_with_buffer': hauler['trip_started_with_buffer'],
                    'fuel': round(fuel_consumption, 2),
                    'lift': round(lift, 2),
                    'avg_tonnes': 0,
                    'started_in_maintenance': started_in_maint,
                    'started_in_maintenance_km': round(started_in_maint_km, 2),
                    'dump_points_count': len(dump_points),
                    'time_at_excavator': round(lead1_time, 1),
                    'wait_at_excavator': round(ex_wait, 1),
                    'time_at_dump': round(lead2_time, 1),
                    'wait_at_dump': round(dump_wait, 1),
                    'total_trip_time': round(total_trip_time, 1),
                    'path_deviation': path_deviation,
                    'path_duration': round(offroute_time, 1),
                    'points': hauler['current_trip_points'].copy()
                }
                
                hauler['total_loaded_distance_km'] += lead_distance_km
                hauler['total_trip_distance_km'] += cycle_distance_km
                hauler['total_maintenance_distance_km'] += maintenance_distance_km
                hauler['trips'].append(trip)
                
                print(f"🚚 {device_id}: TRIP END - Maintenance! Lead: {lead_distance_km:.2f}km, Cycle: {cycle_distance_km:.2f}km, Maint: {maintenance_distance_km:.2f}km", file=sys.stderr)
                print(f"   Lead1: {lead1_time:.1f}min, Lead2: {lead2_time:.1f}min, Total: {total_trip_time:.1f}min", file=sys.stderr)
                print(f"   Wait Ex: {ex_wait:.1f}min, Wait Dump: {dump_wait:.1f}min", file=sys.stderr)
                print(f"   Path Deviation: {path_deviation}, Path Duration: {offroute_time:.1f}min", file=sys.stderr)
                
                hauler['current_trip'] = None
                hauler['dump_start_time'] = None
                hauler['trip_points'] = []
                hauler['maintenance_points'] = []
                hauler['last_trip_point'] = None
                hauler['trip_accumulated_distance'] = 0
                hauler['excavator_start_lat'] = None
                hauler['excavator_start_lon'] = None
                hauler['trip_unknown_points'] = []
                hauler['trip_started_with_buffer'] = False
                hauler['trip_first_excavator_distance'] = 0
                hauler['trip_first_dump_distance'] = 0
                hauler['trip_last_dump_distance'] = 0
                hauler['trip_end_distance'] = 0
                hauler['first_dump_found'] = False
                hauler['maintenance_to_excavator_distance'] = 0
                hauler['current_trip_max_altitude'] = 0
                hauler['trip_start_fuel_raw'] = None
                hauler['last_fuel_level'] = 0
                hauler['trip_dump_points'] = []
                hauler['excavator_return_point'] = None
                hauler['offroute_time'] = 0.0
            
            if hauler['current_trip'] is None and loc_type != 'maintenance' and loc_type != 'excavator' and loc_type != 'dump':
                hauler['in_trip'] = False
            
            # ============================================
            # UPDATE PREVIOUS VALUES - MOVED TO THE END!
            # ============================================
            hauler['prev_timestamp'] = timestamp
            hauler['prev_location_type'] = loc_type
            hauler['prev_lat'] = lat
            hauler['prev_lon'] = lon
            hauler['prev_is_moving'] = is_moving
            hauler['prev_is_engine_on'] = is_engine_on
            hauler['prev_on_route'] = on_route  # ADDED: Store previous point's route status
        
        # ============================================
        # HANDLE INCOMPLETE MAINTENANCE AT END OF DATA
        # ============================================
        if hauler['was_in_maintenance_prev'] and hauler['maintenance_entry_time'] is not None:
            duration = calculate_duration(hauler['maintenance_entry_time'], hauler['last_record'])
            if duration > 0:
                hauler['total_maintenance_minutes'] += duration
                hauler['maintenance_periods'].append({
                    'entry': hauler['maintenance_entry_time'],
                    'exit': hauler['last_record'],
                    'duration_minutes': duration,
                    'incomplete': True
                })
                print(f"🔧 {device_id}: MAINTENANCE PERIOD ENDED AT DATA END - Duration: {duration:.1f} min (incomplete)", file=sys.stderr)
            
            hauler['maintenance_entry_time'] = None
            hauler['maintenance_entry_lat'] = None
            hauler['maintenance_entry_lon'] = None
            hauler['was_in_maintenance_prev'] = False
        
        # ============================================
        # HANDLE INCOMPLETE TRIPS
        # ============================================
        if hauler['current_trip'] is not None and hauler['at_dump']:
            hauler['trip_ended_in_unknown'] = True
            
            # ===== ADD THIS LINE =====
            trip_points = hauler.get('current_trip_points', [])
            # ===== END =====
            
            dump_points = hauler['trip_dump_points']
            
            if dump_points:
                dump_points.sort(key=lambda x: x.get('distance', 0))
                
                if len(dump_points) >= 5:
                    middle_index = len(dump_points) // 2
                    selected_dump = dump_points[middle_index]
                elif len(dump_points) == 2:
                    selected_dump = dump_points[1]
                elif len(dump_points) == 3:
                    selected_dump = dump_points[1]
                else:
                    selected_dump = dump_points[0]
                
                lead_distance_km = selected_dump.get('lead_distance', hauler['trip_first_dump_distance'])
                dump_time = selected_dump.get('timestamp', '')
            else:
                lead_distance_km = hauler['trip_first_dump_distance']
                dump_time = hauler['current_trip'].get('dump_arrival', '')
            
            cycle_distance_km = hauler['trip_accumulated_distance']
            
            lead1_time = calculate_duration(hauler['current_trip']['excavator_arrival'], dump_time)
            lead2_time = calculate_duration(dump_time, hauler['last_record'])
            total_trip_time = calculate_duration(hauler['current_trip']['excavator_arrival'], hauler['last_record'])
            
            ex_wait = hauler.get('excavator_wait_minutes', 0) - hauler.get('trip_start_excavator_wait', 0)
            dump_wait = hauler.get('dump_wait_minutes', 0) - hauler.get('trip_start_dump_wait', 0)
            
            offroute_time = hauler['current_trip'].get('offroute_time', 0)
            path_deviation = 'YES' if offroute_time > 5 else 'NO'
            
            start_fuel_raw = hauler.get('trip_start_fuel_raw')
            end_fuel_raw = hauler.get('last_fuel_level')
            if start_fuel_raw is None:
                start_fuel_raw = 0
            if end_fuel_raw is None:
                end_fuel_raw = 0
            
            if start_fuel_raw > 0 and end_fuel_raw > 0 and start_fuel_raw > end_fuel_raw:
                fuel_consumption = (start_fuel_raw - end_fuel_raw) / 5
            else:
                fuel_consumption = 0
            
            started_in_maint = hauler.get('trip_started_in_maintenance', False)
            started_in_maint_km = hauler.get('trip_started_in_maintenance_distance', 0)
            
            max_alt = hauler.get('current_trip_max_altitude', 0)
            base_alt = hauler.get('lift_base_altitude', 390)
            lift = max_alt - base_alt if max_alt > base_alt else 0
            print(f"📊 {device_id}: LIFT CALC - Max Alt: {max_alt:.1f}m, Base: {base_alt}, Lift: {lift:.2f}m", file=sys.stderr)
            
            trip = {
                'excavator_arrival': hauler['current_trip']['excavator_arrival'],
                'dump_arrival': dump_time,
                'trip_end': hauler['last_record'],
                'dump_location': hauler['trip_dump_location'],
                'loaded_distance_km': round(lead_distance_km, 2),
                'trip_distance_km': round(cycle_distance_km, 2),
                'ended_in_maintenance': False,
                'ended_in_unknown': True,
                'maintenance_distance_km': 0,
                'visited_maintenance': hauler['trip_visited_maintenance'],
                'started_with_buffer': hauler['trip_started_with_buffer'],
                'fuel': round(fuel_consumption, 2),
                'lift': round(lift, 2),
                'avg_tonnes': 0,
                'started_in_maintenance': started_in_maint,
                'started_in_maintenance_km': round(started_in_maint_km, 2),
                'dump_points_count': len(dump_points),
                'time_at_excavator': round(lead1_time, 1),
                'wait_at_excavator': round(ex_wait, 1),
                'time_at_dump': round(lead2_time, 1),
                'wait_at_dump': round(dump_wait, 1),
                'total_trip_time': round(total_trip_time, 1),
                'path_deviation': path_deviation,
                'path_duration': round(offroute_time, 1),
                'points': hauler['current_trip_points'].copy()
            }
            
            hauler['total_loaded_distance_km'] += lead_distance_km
            hauler['total_trip_distance_km'] += cycle_distance_km
            hauler['trips'].append(trip)
            
            print(f"🚚 {device_id}: TRIP END - Unknown (incomplete). Lead: {lead_distance_km:.2f}km, Cycle: {cycle_distance_km:.2f}km", file=sys.stderr)
            print(f"   Lead1: {lead1_time:.1f}min, Lead2: {lead2_time:.1f}min, Total: {total_trip_time:.1f}min", file=sys.stderr)
            print(f"   Wait Ex: {ex_wait:.1f}min, Wait Dump: {dump_wait:.1f}min", file=sys.stderr)
            print(f"   Path Deviation: {path_deviation}, Path Duration: {offroute_time:.1f}min", file=sys.stderr)
            
            hauler['current_trip'] = None
            hauler['dump_start_time'] = None
            hauler['trip_points'] = []
            hauler['maintenance_points'] = []
            hauler['last_trip_point'] = None
            hauler['trip_accumulated_distance'] = 0
            hauler['excavator_start_lat'] = None
            hauler['excavator_start_lon'] = None
            hauler['trip_unknown_points'] = []
            hauler['trip_started_with_buffer'] = False
            hauler['trip_first_excavator_distance'] = 0
            hauler['trip_first_dump_distance'] = 0
            hauler['trip_last_dump_distance'] = 0
            hauler['trip_end_distance'] = 0
            hauler['first_dump_found'] = False
            hauler['maintenance_to_excavator_distance'] = 0
            hauler['current_trip_max_altitude'] = 0
            hauler['trip_start_fuel_raw'] = None
            hauler['last_fuel_level'] = 0
            hauler['trip_dump_points'] = []
            hauler['excavator_return_point'] = None
            hauler['offroute_time'] = 0.0
        
        # ============================================
        # CALCULATE AVERAGE LIFT
        # ============================================
        if hauler['trips']:
            total_lift = sum(trip.get('lift', 0) for trip in hauler['trips'])
            hauler['avg_lift'] = total_lift / len(hauler['trips'])
        else:
            hauler['avg_lift'] = 0
        
        print(f"\n📊 {device_id} (Hauler {hauler['hauler']}) Summary:", file=sys.stderr)
        print(f"   Productive Idle (Ex+Dump): {hauler['productive_idle_minutes']:.1f} min", file=sys.stderr)
        print(f"     - Engine ON: {hauler['productive_idle_engine_on_minutes']:.1f} min", file=sys.stderr)
        print(f"     - Engine OFF: {hauler['productive_idle_engine_off_minutes']:.1f} min", file=sys.stderr)
        print(f"   Non-Productive Idle (On-Route): {hauler['nonproductive_onroute_idle_minutes']:.1f} min", file=sys.stderr)
        print(f"     - Engine ON: {hauler['nonproductive_onroute_engine_on_minutes']:.1f} min", file=sys.stderr)
        print(f"     - Engine OFF: {hauler['nonproductive_onroute_engine_off_minutes']:.1f} min", file=sys.stderr)
        print(f"   Non-Productive Idle (Off-Route): {hauler['nonproductive_offroute_idle_minutes']:.1f} min", file=sys.stderr)
        print(f"     - Engine ON: {hauler['nonproductive_offroute_engine_on_minutes']:.1f} min", file=sys.stderr)
        print(f"     - Engine OFF: {hauler['nonproductive_offroute_engine_off_minutes']:.1f} min", file=sys.stderr)
        print(f"   Travel Time: {hauler['travel_time_minutes']:.1f} min", file=sys.stderr)
        print(f"   Maintenance: {hauler['total_maintenance_minutes']:.1f} min", file=sys.stderr)
        print(f"   Average Lift: {hauler['avg_lift']:.2f} m", file=sys.stderr)
        
        total_shift = (hauler['productive_idle_minutes'] + 
                       hauler['nonproductive_onroute_idle_minutes'] + 
                       hauler['nonproductive_offroute_idle_minutes'] + 
                       hauler['travel_time_minutes'] + 
                       hauler['total_maintenance_minutes'])
        print(f"   Total Shift: {total_shift:.1f} min", file=sys.stderr)
        print(f"   Expected Shift Duration: {calculate_duration(hauler['first_record'], hauler['last_record']):.1f} min", file=sys.stderr)
    
    return trips_data, all_point_data


# ============================================
# MAP GENERATION FUNCTIONS
# ============================================

def generate_trip_map(hauler_number, trip_number, trip_points, dump_location, output_dir):
    if not FOLIUM_AVAILABLE or not trip_points:
        return None, None
    
    # Convert points to consistent format
    formatted_points = []
    for p in trip_points:
        if isinstance(p, dict):
            formatted_points.append(p)
        elif isinstance(p, tuple) and len(p) >= 3:
            formatted_points.append({
                'lat': p[0],
                'lon': p[1],
                'time': p[2] if len(p) > 2 else ''
            })
        else:
            continue
    
    if not formatted_points:
        return None, None
    
    coordinates = [(p['lat'], p['lon']) for p in formatted_points if 'lat' in p and 'lon' in p]
    if not coordinates:
        return None, None
    
    center_lat = coordinates[0][0]
    center_lon = coordinates[0][1]
    
    color_idx = (trip_number - 1) % len(TRIP_COLORS)
    color = TRIP_COLORS[color_idx]
    
    map_obj = folium.Map(location=[center_lat, center_lon], zoom_start=15, control_scale=True)
    
    # Simple trip path
    folium.PolyLine(coordinates, color=color, weight=4, opacity=0.9).add_to(map_obj)
    
    # Simple markers
    if coordinates:
        folium.Marker(coordinates[0], icon=folium.Icon(color='green', icon='play', prefix='fa')).add_to(map_obj)
        folium.Marker(coordinates[-1], icon=folium.Icon(color='red', icon='stop', prefix='fa')).add_to(map_obj)
    
    # Simple dump locations (no labels)
    for dump in DUMP_LOCATIONS:
        folium.CircleMarker(
            location=[dump['lat'], dump['lon']],
            radius=8, color='orange', fill=True, fill_color='orange', fill_opacity=0.6
        ).add_to(map_obj)
    
    # Polygon dump
    if POLYGON_DUMP:
        poly_coords = [(c[0], c[1]) for c in POLYGON_DUMP['coordinates']]
        folium.Polygon(locations=poly_coords, color='orange', weight=2, fill=True, fill_color='orange', fill_opacity=0.15).add_to(map_obj)
    
    # Excavator polygon
    ex_coords = [(c[0], c[1]) for c in EXCAVATOR_POLYGON['coordinates']]
    folium.Polygon(locations=ex_coords, color='green', weight=2, fill=True, fill_color='green', fill_opacity=0.15).add_to(map_obj)
    
    # Maintenance circle
    folium.Circle(
        location=[MAINTENANCE_LAT, MAINTENANCE_LON],
        radius=MAINTENANCE_RADIUS, color='red', fill=True, fill_color='red', fill_opacity=0.1
    ).add_to(map_obj)
    
    # REMOVE LEGEND - THIS SAVES A LOT OF SPACE!
    
    html_filename = f"hauler_{hauler_number}_trip_{trip_number}.html"
    html_path = os.path.join(output_dir, html_filename)
    map_obj.save(html_path)
    
    # Capture image with smaller size
    image_path = None
    if SELENIUM_AVAILABLE:
        try:
            chrome_options = Options()
            chrome_options.add_argument("--headless")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--window-size=600,400")  # SMALLER!
            chrome_options.add_argument("--disable-gpu")
            
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=chrome_options)
            driver.get(f"file://{html_path}")
            time.sleep(2)
            
            image_filename = f"hauler_{hauler_number}_trip_{trip_number}.png"
            image_path = os.path.join(output_dir, image_filename)
            driver.save_screenshot(image_path)
            driver.quit()
        except Exception as e:
            image_path = None
    
    return html_path, image_path


def create_excel(trips_data, all_point_data, excavator_results=None):
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)
    
    # ============================================
    # XLSXWRITER FORMATS
    # ============================================
    header_format = workbook.add_format({
        'bold': True,
        'color': 'white',
        'bg_color': '#4472C4',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })
    
    center_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })
    
    title_format = workbook.add_format({
        'bold': True,
        'font_size': 14,
        'align': 'center',
        'valign': 'vcenter'
    })
    
    subtitle_format = workbook.add_format({
        'bold': True,
        'font_size': 12,
        'align': 'center',
        'valign': 'vcenter'
    })
    
    wrap_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True
    })
    
    bold_center = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter'
    })
    
    temp_dir = tempfile.mkdtemp(prefix="trip_maps_")
    shift_label = detect_shift_from_data(trips_data)
    
    # ============================================
    # HAULER SUMMARY SHEET - YOUR EXACT 17 COLUMNS (A-Q)
    # ============================================
    ws_summary = workbook.add_worksheet("Hauler Summary")
    
    ws_summary.merge_range('A1:Q1', 'MINING HAULER - SHIFT ANALYSIS', title_format)
    ws_summary.merge_range('A2:Q2', f'Shift: {shift_label}', subtitle_format)
    ws_summary.merge_range('A8:Q8', 'Time in minutes', bold_center)
    ws_summary.merge_range('A9:Q9', 'Distance in KM', bold_center)
    
    # YOUR EXACT 17 COLUMN HEADERS
    headers = ['Hauler', 'Trips', 
               'Lead Distance', 'Cycle Distance', 'Total Distance',
               'Avg Cycle Time', 'Running Time', 
               'Productive Idle', 'Non-Prod On-Route', 'Non-Prod Off-Route',
               'Travel Time', 'In Maintenance', 
               'Fuel(L)', 'AVG Lift', 'Avg L/Hr', 'Avg L/Tonne', 'Tonne/Hr']
    
    row = 3
    for col, header in enumerate(headers):
        ws_summary.write(row, col, header, header_format)
        ws_summary.set_column(col, col, 16)
    
    row = 4
    for hauler in HAULER_SHEETS:
        found = False
        device_data_found = None
        
        for device_id, device_data in trips_data.items():
            if device_data.get('hauler') == hauler:
                found = True
                device_data_found = device_data
                break
        
        if found:
            trips = device_data_found.get('trips', [])
            
            # Check if hauler has any GPS points
            has_gps_data = False
            for point in all_point_data:
                if point.get('hauler') == hauler:
                    has_gps_data = True
                    break
            
            # Show hauler if it has trips OR has GPS data
            if len(trips) > 0 or has_gps_data:
                loaded_dist = device_data_found.get('total_loaded_distance_km', 0)
                trip_dist = device_data_found.get('total_trip_distance_km', 0)
                maint_dist = device_data_found.get('total_maintenance_distance_km', 0)
                
                # Total Distance = Cycle Distance + Maintenance Distance
                started_maint_total = 0
                for trip in trips:
                    if trip.get('started_in_maintenance', False):
                        started_maint_total += trip.get('started_in_maintenance_km', 0)
                
                total_dist = trip_dist + started_maint_total
                
                running_time = sum(t.get('total_trip_time', 0) for t in trips)
                avg_cycle_time = running_time / len(trips) if len(trips) > 0 else 0
                
                productive_idle = device_data_found.get('productive_idle_minutes', 0)
                nonproductive_onroute = device_data_found.get('nonproductive_onroute_idle_minutes', 0)
                nonproductive_offroute = device_data_found.get('nonproductive_offroute_idle_minutes', 0)
                travel_time = device_data_found.get('travel_time_minutes', 0)
                in_maintenance = device_data_found.get('total_maintenance_minutes', 0)
                
                fuel = sum(t.get('fuel', 0) for t in trips)
                avg_lift = device_data_found.get('avg_lift', 0)
                
                engine_on_time_minutes = running_time + nonproductive_onroute + nonproductive_offroute
                engine_on_time_hours = engine_on_time_minutes / 60
                avg_l_per_hr = fuel / engine_on_time_hours if engine_on_time_hours > 0 else 0
                
                total_tonnes = len(trips) * 35
                avg_l_per_tonne = fuel / total_tonnes if total_tonnes > 0 else 0
                
                running_time_hours = running_time / 60
                tonne_per_hr = total_tonnes / running_time_hours if running_time_hours > 0 else 0
                
                # Use different format for hauler with data but no trips
                if len(trips) == 0:
                    highlight_format = workbook.add_format({
                        'border': 1,
                        'align': 'center',
                        'valign': 'vcenter',
                        'bg_color': '#FFE5B4',
                        'font_color': '#CC6600'
                    })
                    fmt = highlight_format
                else:
                    fmt = center_format
                
                ws_summary.write(row, 0, f"Hauler {hauler}", fmt)
                ws_summary.write(row, 1, len(trips), fmt)
                ws_summary.write(row, 2, round(loaded_dist, 2), fmt)
                ws_summary.write(row, 3, round(trip_dist, 2), fmt)
                ws_summary.write(row, 4, round(total_dist, 2), fmt)
                ws_summary.write(row, 5, round(avg_cycle_time, 2), fmt)
                ws_summary.write(row, 6, round(running_time, 2), fmt)
                ws_summary.write(row, 7, round(productive_idle, 2), fmt)
                ws_summary.write(row, 8, round(nonproductive_onroute, 2), fmt)
                ws_summary.write(row, 9, round(nonproductive_offroute, 2), fmt)
                ws_summary.write(row, 10, round(travel_time, 2), fmt)
                ws_summary.write(row, 11, round(in_maintenance, 2), fmt)
                ws_summary.write(row, 12, round(fuel, 2), fmt)
                ws_summary.write(row, 13, round(avg_lift, 2), fmt)
                ws_summary.write(row, 14, round(avg_l_per_hr, 2), fmt)
                ws_summary.write(row, 15, round(avg_l_per_tonne, 2), fmt)
                ws_summary.write(row, 16, round(tonne_per_hr, 2), fmt)
                row += 1
    
    # ============================================
    # IDLE TIME ANALYSIS - YOUR EXACT 10 COLUMNS (A-J)
    # ============================================
    ws_idle = workbook.add_worksheet("Idle Time Analysis")
    ws_idle.merge_range('A1:J1', 'ENGINE ON/OFF IDLE TIME BREAKDOWN', title_format)
    
    idle_headers = ['Hauler', 
                    'Prod Idle (ON)', 'Prod Idle (OFF)', 'Prod Idle Total',
                    'On-Route (ON)', 'On-Route (OFF)', 'On-Route Total',
                    'Off-Route (ON)', 'Off-Route (OFF)', 'Off-Route Total']
    
    row = 2
    for col, header in enumerate(idle_headers):
        ws_idle.write(row, col, header, header_format)
        ws_idle.set_column(col, col, 16)
    
    row = 3
    for hauler in HAULER_SHEETS:
        found = False
        device_data_found = None
        
        for device_id, device_data in trips_data.items():
            if device_data.get('hauler') == hauler:
                found = True
                device_data_found = device_data
                break
        
        if found:
            trips = device_data_found.get('trips', [])
            
            # Check if hauler has any GPS points
            has_gps_data = False
            for point in all_point_data:
                if point.get('hauler') == hauler:
                    has_gps_data = True
                    break
            
            # Show hauler if it has trips OR has GPS data
            if len(trips) > 0 or has_gps_data:
                # Use different format for hauler with data but no trips
                if len(trips) == 0:
                    highlight_format = workbook.add_format({
                        'border': 1,
                        'align': 'center',
                        'valign': 'vcenter',
                        'bg_color': '#FFE5B4',
                        'font_color': '#CC6600'
                    })
                    fmt = highlight_format
                else:
                    fmt = center_format
                
                ws_idle.write(row, 0, f"Hauler {hauler}", fmt)
                ws_idle.write(row, 1, round(device_data_found.get('productive_idle_engine_on_minutes', 0), 1), fmt)
                ws_idle.write(row, 2, round(device_data_found.get('productive_idle_engine_off_minutes', 0), 1), fmt)
                ws_idle.write(row, 3, round(device_data_found.get('productive_idle_minutes', 0), 1), fmt)
                ws_idle.write(row, 4, round(device_data_found.get('nonproductive_onroute_engine_on_minutes', 0), 1), fmt)
                ws_idle.write(row, 5, round(device_data_found.get('nonproductive_onroute_engine_off_minutes', 0), 1), fmt)
                ws_idle.write(row, 6, round(device_data_found.get('nonproductive_onroute_idle_minutes', 0), 1), fmt)
                ws_idle.write(row, 7, round(device_data_found.get('nonproductive_offroute_engine_on_minutes', 0), 1), fmt)
                ws_idle.write(row, 8, round(device_data_found.get('nonproductive_offroute_engine_off_minutes', 0), 1), fmt)
                ws_idle.write(row, 9, round(device_data_found.get('nonproductive_offroute_idle_minutes', 0), 1), fmt)
                row += 1
    
    # ============================================
    # TRIP SHEETS - YOUR EXACT 18 COLUMNS (A-R) WITH MAPS
    # ============================================
    for hauler in HAULER_SHEETS:
        hauler_found = False
        hauler_data = None
        
        for device_id, device_data in trips_data.items():
            if device_data.get('hauler') == hauler:
                hauler_found = True
                hauler_data = device_data
                break
        
        # Skip only if hauler has NO data at all
        if not hauler_found:
            # Check if hauler has any GPS points
            has_gps_data = False
            for point in all_point_data:
                if point.get('hauler') == hauler:
                    has_gps_data = True
                    break
            
            if not has_gps_data:
                continue  # Skip haulers with no data at all
            else:
                # Create dummy data for hauler with GPS data but not in trips_data
                hauler_data = {
                    'hauler': hauler,
                    'trips': [],
                    'total_maintenance_minutes': 0,
                    'productive_idle_minutes': 0,
                    'nonproductive_onroute_idle_minutes': 0,
                    'nonproductive_offroute_idle_minutes': 0,
                    'travel_time_minutes': 0,
                }
                trips_data[hauler] = hauler_data
        
        trips = hauler_data.get('trips', [])
        
        # Check if hauler has any GPS points
        has_gps_data = False
        for point in all_point_data:
            if point.get('hauler') == hauler:
                has_gps_data = True
                break
        
        # Skip only if NO data at all (no trips AND no GPS data)
        if len(trips) == 0 and not has_gps_data:
            continue
        
        sheet_name = f"Hauler {hauler} Trips"
        ws_trip = workbook.add_worksheet(sheet_name)
        
        ws_trip.merge_range('A1:R1', f'Hauler {hauler} - Trip Details', title_format)
        
        # YOUR EXACT 18 COLUMN TRIP HEADERS
        trip_headers = ['Trip No', 'Dump Location', 'Lead Distance (km)', 'Cycle Distance (km)',
                       'Lead1(min)', 'Wait at Ex', 'Lead2(min)', 'Wait at Dump', 'Cycle Time(min)', 
                       'Path Deviation', 'Path Duration (min)',
                       'Lift', 'Avg tonnes',
                       'Started Maint(km)', 'Ended Maint (km)', 
                       'Started in Maint', 'Ended in Maintenance', 'Ended in Unknown']
        
        row = 3
        for col, header in enumerate(trip_headers):
            ws_trip.write(row, col, header, header_format)
            ws_trip.set_column(col, col, 16)
        
        row = 4
        
        # If no trips but has GPS data, show message
        if len(trips) == 0:
            maint_minutes = hauler_data.get('total_maintenance_minutes', 0)
            if maint_minutes > 0:
                ws_trip.merge_range(row, 0, row, 17, 
                                   f"⚠️ No trips recorded - Hauler spent {maint_minutes:.1f} minutes in maintenance", 
                                   workbook.add_format({
                                       'bold': True,
                                       'font_size': 12,
                                       'align': 'center',
                                       'valign': 'vcenter',
                                       'font_color': '#CC6600',
                                       'bg_color': '#FFE5B4',
                                       'border': 1
                                   }))
            else:
                ws_trip.merge_range(row, 0, row, 17, 
                                   "⚠️ No trips recorded - Hauler had GPS activity but no complete trips", 
                                   workbook.add_format({
                                       'bold': True,
                                       'font_size': 12,
                                       'align': 'center',
                                       'valign': 'vcenter',
                                       'font_color': '#CC6600',
                                       'bg_color': '#FFE5B4',
                                       'border': 1
                                   }))
            row += 1
            continue  # Skip map generation for this hauler
        
        trip_image_data = []
        
        for i, trip in enumerate(trips, 1):
            trip_number = trip.get('trip_number', i)
            
            ws_trip.write(row, 0, trip_number, center_format)
            ws_trip.write(row, 1, trip.get('dump_location', 'Unknown'), wrap_format)
            ws_trip.write(row, 2, trip.get('loaded_distance_km', 0), center_format)
            ws_trip.write(row, 3, trip.get('trip_distance_km', 0), center_format)
            ws_trip.write(row, 4, trip.get('time_at_excavator', 0), center_format)
            ws_trip.write(row, 5, trip.get('wait_at_excavator', 0), center_format)
            ws_trip.write(row, 6, trip.get('time_at_dump', 0), center_format)
            ws_trip.write(row, 7, trip.get('wait_at_dump', 0), center_format)
            ws_trip.write(row, 8, trip.get('total_trip_time', 0), center_format)
            
            # ===== PATH DEVIATION - Column 10 =====
            path_dev = trip.get('path_deviation', 'NO')
            ws_trip.write(row, 9, path_dev, center_format)
            
            # ===== PATH DURATION - Column 11 (index 10) =====
            path_duration = trip.get('path_duration', 0)
            if path_duration > 5:
                ws_trip.write(row, 10, round(path_duration, 1), center_format)
            else:
                ws_trip.write(row, 10, '0', center_format)
                
            ws_trip.write(row, 11, round(trip.get('lift', 0), 2), center_format)
            ws_trip.write(row, 12, round(trip.get('avg_tonnes', 0), 2), center_format)
            
            started_maint_km = trip.get('started_in_maintenance_km', 0)
            ws_trip.write(row, 13, started_maint_km if started_maint_km > 0 else '', center_format)
            
            ended_maint_km = trip.get('maintenance_distance_km', 0)
            ws_trip.write(row, 14, ended_maint_km if ended_maint_km > 0 else '', center_format)
            
            started_in_maint = trip.get('started_in_maintenance', False)
            ws_trip.write(row, 15, "YES" if started_in_maint else "NO", center_format)
            
            ended_in_maint = trip.get('ended_in_maintenance', False)
            ws_trip.write(row, 16, "YES" if ended_in_maint else "NO", center_format)
            
            ended_in_unknown = trip.get('ended_in_unknown', False)
            ws_trip.write(row, 17, "YES" if ended_in_unknown else "NO", center_format)
            
            # Generate map for this trip
            trip_points = trip.get('points', [])
            if FOLIUM_AVAILABLE and trip_points:
                try:
                    html_path, image_path = generate_trip_map(
                        hauler, trip_number, trip_points, trip.get('dump_location', 'Unknown'), temp_dir
                    )
                    if image_path and os.path.exists(image_path):
                        trip_image_data.append((trip_number, image_path, trip))
                        print(f"✅ Map for Hauler {hauler} Trip {trip_number} captured", file=sys.stderr)
                except Exception as e:
                    pass
            
            row += 1
                        
                # Map generation disabled
            trip_points = trip.get('points', [])
                # Skip map generation - screenshots disabled         
    # ============================================
    # POINT ANALYSIS - YOUR EXACT 17 COLUMNS (A-Q)
    # ============================================
    ws_points = workbook.add_worksheet("Point-by-Point Analysis")
    
    ws_points.merge_range('A1:Q1', 'POINT-BY-POINT GPS ANALYSIS WITH ROUTE STATUS', title_format)
    ws_points.merge_range(2, 0, 2, 16, f"Total Points: {len(all_point_data)}", bold_center)
    ws_points.merge_range(3, 0, 3, 16, f"Excavator Polygon: {len(EXCAVATOR_POLYGON['coordinates'])} points, Buffer: {EXCAVATOR_BUFFER_METERS}m", bold_center)
    ws_points.merge_range(4, 0, 4, 16, f"GPX Routes Loaded: {len(GPX_FILE_PATHS)} files, Buffer: {ROUTE_BUFFER_METERS}m", bold_center)
    
    point_headers = ['Hauler', 'Timestamp', 'Latitude', 'Longitude', 'Location Type', 'Location Name',
                     'Route Status', 'Distance (km)', 'Fuel Consumption (L)', 'Pitch (°)', 'Roll (°)', 
                     'Vibration (m/s²)', 'Fuel (%)', 'Speed (km/h)', 'Altitude (m)',
                     'Is Moving', 'Movement (m)']
    
    row = 6
    for col, header in enumerate(point_headers):
        ws_points.write(row, col, header, header_format)
        ws_points.set_column(col, col, 16)
    
    row = 7
    for point in all_point_data:
        ws_points.write(row, 0, f"Hauler {point['hauler']}", center_format)
        ws_points.write(row, 1, point['timestamp'], center_format)
        ws_points.write(row, 2, round(point['lat'], 6), center_format)
        ws_points.write(row, 3, round(point['lon'], 6), center_format)
        ws_points.write(row, 4, point['location_type'], center_format)
        ws_points.write(row, 5, point['location_name'], center_format)
        ws_points.write(row, 6, point.get('route_status', 'Off Route'), center_format)
        ws_points.write(row, 7, round(point.get('distance', 0), 2), center_format)
        ws_points.write(row, 8, round(point.get('fuel_consumption', 0), 2), center_format)
        ws_points.write(row, 9, round(point.get('pitch', 0), 1), center_format)
        ws_points.write(row, 10, round(point.get('roll', 0), 1), center_format)
        ws_points.write(row, 11, round(point.get('vibration', 0), 3), center_format)
        ws_points.write(row, 12, round(point.get('fuel', 0), 1), center_format)
        ws_points.write(row, 13, round(point.get('speed', 0), 1), center_format)
        ws_points.write(row, 14, round(point.get('altitude', 0), 2), center_format)
        ws_points.write(row, 15, "YES" if point.get('is_moving', False) else "NO", center_format)
        ws_points.write(row, 16, round(point.get('movement_distance_m', 0), 1), center_format)
        row += 1
    
    # Summary
    row += 2
    ws_points.merge_range(row, 0, row, 16, 'POINT ANALYSIS SUMMARY', title_format)
    row += 1
    
    loc_counts = {}
    route_counts = {'On Route': 0, 'Off Route': 0}
    
    for point in all_point_data:
        loc_type = point['location_type']
        loc_counts[loc_type] = loc_counts.get(loc_type, 0) + 1
        route_status = point.get('route_status', 'Off Route')
        route_counts[route_status] = route_counts.get(route_status, 0) + 1
    
    ws_points.write(row, 0, "Location Type Distribution:", bold_center)
    row += 1
    
    for loc_type, count in loc_counts.items():
        ws_points.write(row, 0, f"  {loc_type.upper()}:", center_format)
        ws_points.write(row, 1, count, center_format)
        pct = (count / len(all_point_data)) * 100 if all_point_data else 0
        ws_points.write(row, 2, f"{pct:.1f}%", center_format)
        row += 1
    
    row += 1
    ws_points.write(row, 0, "Route Status Distribution:", bold_center)
    row += 1
    
    for status, count in route_counts.items():
        ws_points.write(row, 0, f"  {status}:", center_format)
        ws_points.write(row, 1, count, center_format)
        pct = (count / len(all_point_data)) * 100 if all_point_data else 0
        ws_points.write(row, 2, f"{pct:.1f}%", center_format)
        row += 1
    
    # ============================================
    # TERRAIN ANALYSIS - YOUR EXACT 7 COLUMNS (A-G)
    # ============================================
    ws_terrain = workbook.add_worksheet("Terrain Analysis")
    
    ws_terrain.merge_range('A1:G1', 'TERRAIN ANALYSIS - PITCH & ROLL', title_format)
    
    terrain_headers = ['Hauler', 'Steep UP', 'Steep Down', 'Flat', 'Max Pitch', 'Min Pitch', 'Avg Pitch']
    
    row = 2
    for col, header in enumerate(terrain_headers):
        ws_terrain.write(row, col, header, header_format)
        ws_terrain.set_column(col, col, 16)
    
    row = 3
    for hauler in HAULER_SHEETS:
        hauler_exists = False
        for device_id, device_data in trips_data.items():
            if device_data.get('hauler') == hauler:
                trips = device_data.get('trips', [])
                if len(trips) > 0:
                    hauler_exists = True
                break
        
        if not hauler_exists:
            # Check if hauler has any GPS points
            has_gps_data = False
            for point in all_point_data:
                if point.get('hauler') == hauler:
                    has_gps_data = True
                    break
            
            if not has_gps_data:
                continue
        
        terrain_data = {'steep_up': 0, 'steep_down': 0, 'flat': 0, 
                       'max_pitch': -999, 'min_pitch': 999, 'pitch_values': []}
        
        for point in all_point_data:
            if point.get('hauler') != hauler:
                continue
            distance = point.get('distance', 0)
            pitch = point.get('pitch', 0)
            if distance > 0 and distance <= 5.0:
                if pitch > 8:
                    terrain_data['steep_up'] += distance
                elif pitch < -8:
                    terrain_data['steep_down'] += distance
                else:
                    terrain_data['flat'] += distance
                if pitch != 0:
                    if pitch > terrain_data['max_pitch']:
                        terrain_data['max_pitch'] = pitch
                    if pitch < terrain_data['min_pitch']:
                        terrain_data['min_pitch'] = pitch
                    terrain_data['pitch_values'].append(pitch)
        
        avg_pitch = sum(terrain_data['pitch_values']) / len(terrain_data['pitch_values']) if terrain_data['pitch_values'] else 0
        if terrain_data['max_pitch'] == -999:
            terrain_data['max_pitch'] = 0
        if terrain_data['min_pitch'] == 999:
            terrain_data['min_pitch'] = 0
        
        ws_terrain.write(row, 0, f"Hauler {hauler}", center_format)
        ws_terrain.write(row, 1, round(terrain_data['steep_up'], 2), center_format)
        ws_terrain.write(row, 2, round(terrain_data['steep_down'], 2), center_format)
        ws_terrain.write(row, 3, round(terrain_data['flat'], 2), center_format)
        ws_terrain.write(row, 4, round(terrain_data['max_pitch'], 1), center_format)
        ws_terrain.write(row, 5, round(terrain_data['min_pitch'], 1), center_format)
        ws_terrain.write(row, 6, round(avg_pitch, 1), center_format)
        row += 1
    
    # Notes
    note_row = row + 1
    ws_terrain.merge_range(note_row, 0, note_row, 6, "Note: Terrain classification based on pitch values", bold_center)
    note_row += 1
    ws_terrain.merge_range(note_row, 0, note_row, 6, " Steep UP: Pitch > 8° ", center_format)
    note_row += 1
    ws_terrain.merge_range(note_row, 0, note_row, 6, " Steep Down: Pitch < -8° ", center_format)
    note_row += 1
    ws_terrain.merge_range(note_row, 0, note_row, 6, " Flat: -8° ≤ Pitch ≤ 8° ", center_format)
    
    # ============================================
    # EXCAVATOR ANALYSIS - YOUR EXACT 6 COLUMNS (A-F)
    # ============================================
    ws_excavator = workbook.add_worksheet("Excavator Analysis")
    
    ws_excavator.merge_range('A1:G1', 'EXCAVATOR PERFORMANCE ANALYSIS', title_format)
    ws_excavator.merge_range('A2:G2', f'Shift: {shift_label}', subtitle_format)
    
    excavator_headers = ['Excavator ID', 'Distance (km)', 'Idle %', 'Runtime %', 'Maintenance %', 'Status']
    
    row = 3
    for col, header in enumerate(excavator_headers):
        ws_excavator.write(row, col, header, header_format)
        ws_excavator.set_column(col, col, 18)
    
    row = 4
    if excavator_results:
        for device_id, data in excavator_results.items():
            ws_excavator.write(row, 0, data['equipment_id'], center_format)
            ws_excavator.write(row, 1, data['total_distance'], center_format)
            ws_excavator.write(row, 2, data['idle_percentage'], center_format)
            ws_excavator.write(row, 3, data['runtime_percentage'], center_format)
            ws_excavator.write(row, 4, data['maintenance_percentage'], center_format)
            ws_excavator.write(row, 5, data['operating_status'], center_format)
            row += 1
    else:
        ws_excavator.write(row, 0, "No excavator data available", center_format)
    
    # Note
    row += 2
    ws_excavator.merge_range(row, 0, row, 6, "Note: Excavator analysis uses vibration data (0 = idle) and GPS proximity to maintenance area", 
                            workbook.add_format({'italic': True, 'font_size': 10}))
    
    workbook.close()
    output.seek(0)
    return output, temp_dir
# ============================================
# MAIN
# ============================================

def main():
    try:
        input_data = sys.stdin.read()
        if not input_data:
            print(json.dumps({'status': 'error', 'error': 'No input data'}))
            return
        
        data = json.loads(input_data)
        records = data.get('data', [])
        
        # ===== ADD THIS BLOCK =====
        global ROUTE_POINTS
        ROUTE_POINTS = load_gpx_routes(GPX_FILE_PATHS)
        print(f"🛣️ Total route points loaded: {len(ROUTE_POINTS)}", file=sys.stderr)
        # ===== END ADDED BLOCK =====
        
        print(f"Processing {len(records)} records", file=sys.stderr)
        print(f"GPS jumps > {MAX_GPS_JUMP_KM}km will be ignored", file=sys.stderr)
        
        trips_data, all_point_data = analyze_trips(records)
        excavator_results = analyze_excavator_data(records)
        
        total_trips = sum(len(d.get('trips', [])) for d in trips_data.values())
        print(f"Found {len(trips_data)} haulers with {total_trips} total trips", file=sys.stderr)
        
        excel_buffer, map_dir = create_excel(trips_data, all_point_data, excavator_results)
        excel_base64 = base64.b64encode(excel_buffer.getvalue()).decode('utf-8')
        
        result = {
            'status': 'success',
            'report': excel_base64,
            'filename': 'Trip_Report.xlsx',
            'map_dir': map_dir
        }
        
        print(json.dumps(result))
        print(f"Report generated successfully!", file=sys.stderr)
        print(f"Maps in: {map_dir}", file=sys.stderr)
        
    except Exception as e:
        error_msg = str(e)
        print(f"Error: {error_msg}", file=sys.stderr)
        print(traceback.format_exc(), file=sys.stderr)
        print(json.dumps({'status': 'error', 'error': error_msg}))       
if __name__ == "__main__":
    main()