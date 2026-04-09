# routes/analysis.py - FIXED FUEL CALCULATION BASED ON ACTUAL DATA

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import json
import sys
import base64
from geopy.distance import geodesic
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from sklearn.cluster import DBSCAN
import warnings
warnings.filterwarnings('ignore')

def convert_numpy_types(obj):
    """Convert numpy types to Python native types for JSON serialization"""
    if isinstance(obj, np.integer):
        return int(obj)
    elif isinstance(obj, np.floating):
        if np.isnan(obj) or np.isinf(obj):
            return 0
        return float(obj)
    elif isinstance(obj, np.ndarray):
        return obj.tolist()
    elif isinstance(obj, np.bool_):
        return bool(obj)
    elif isinstance(obj, pd.Timestamp):
        return obj.isoformat()
    elif isinstance(obj, dict):
        return {str(key): convert_numpy_types(value) for key, value in obj.items()}
    elif isinstance(obj, (list, tuple)):
        return [convert_numpy_types(item) for item in obj]
    elif pd.isna(obj):
        return 0
    else:
        return obj

class SimpleMiningAnalytics:
    """Simple Mining Analytics using only essential parameters"""
    
    def __init__(self):
        self.diesel_price = 94.5  # ₹ per liter
        
        # FIXED: Fuel consumption rates based on ACTUAL mining data
        # Base fuel rate on flat terrain: 0.65 L/km
        # Steep Up: 2.5x to 3x more fuel (matches your data showing 1.05 L/km on steep)
        # Steep Down: 0.3x fuel (engine braking, minimal fuel)
        
        self.fuel_rates = {
            'Steep Up': 1.05,      # L/km - matches Device 133 (7.88L / 7.53km = 1.046 L/km)
            'Mild Up': 0.85,       # L/km - moderate incline
            'Flat': 0.65,          # L/km - base consumption (Device 137: 9.44L / 15.03km = 0.63 L/km)
            'Mild Down': 0.45,     # L/km - slight downhill
            'Steep Down': 0.25     # L/km - steep downhill (engine braking)
        }
        
        self.max_realistic_speed = 35  # Max realistic speed for mining hauler (km/h)
        self.max_trips_per_8h = 13  # Maximum realistic trips in 8 hour shift
        
        # Gradient bands based on pitch angle
        self.gradient_bands = {
            'Steep Down': (-float('inf'), -8),   # Less than -8 degrees
            'Mild Down': (-8, -3),               # -8 to -3 degrees
            'Flat': (-3, 3),                     # -3 to 3 degrees
            'Mild Up': (3, 8),                   # 3 to 8 degrees
            'Steep Up': (8, float('inf'))        # Greater than 8 degrees
        }
        
        # TMC Maintenance Area Coordinates
        self.maintenance_lat = 20.4051928
        self.maintenance_lon = 81.0662203
        self.maintenance_radius = 100  # 100 meters radius
        
    def is_in_maintenance_area(self, lat, lon):
        """Check if device is in TMC maintenance area"""
        if lat == 0 or lon == 0:
            return False
        try:
            dist = geodesic((lat, lon), (self.maintenance_lat, self.maintenance_lon)).meters
            return dist <= self.maintenance_radius
        except:
            return False
    
    def calculate_distance_and_fuel(self, df):
        """Calculate distance between points and estimate fuel consumption"""
        distances = []
        cumulative_distance = 0
        cumulative_fuel = 0
        in_maintenance_count = 0
        
        for i in range(len(df)):
            if i == 0:
                distances.append(0)
                continue
                
            p1 = (df.iloc[i-1]['lat'], df.iloc[i-1]['lon'])
            p2 = (df.iloc[i]['lat'], df.iloc[i]['lon'])
            
            # Skip invalid coordinates
            if p1[0] == 0 or p1[1] == 0 or p2[0] == 0 or p2[1] == 0:
                distances.append(0)
                continue
            
            try:
                # Calculate distance in meters
                dist_m = geodesic(p1, p2).meters
                
                # 🚀 Filter out GPS jumps > 5km
                if dist_m > 5000:
                    print(f"  Filtered GPS jump of {dist_m:.0f}m at index {i}", file=sys.stderr)
                    distances.append(0)
                    continue
                
                # Small movements (<1m) are likely GPS noise
                if dist_m < 1:
                    distances.append(0)
                    continue
                    
                distances.append(dist_m)
                dist_km = dist_m / 1000
                cumulative_distance += dist_km
                
                # Calculate fuel based on gradient using FIXED rates
                pitch = df.iloc[i]['pitch']
                gradient_class = self.classify_gradient(pitch)
                fuel_rate = self.fuel_rates.get(gradient_class, 0.65)
                
                # Fuel for this segment (L)
                segment_fuel = dist_km * fuel_rate
                cumulative_fuel += segment_fuel
                
                # Check if in maintenance area
                if self.is_in_maintenance_area(df.iloc[i]['lat'], df.iloc[i]['lon']):
                    in_maintenance_count += 1
                
            except Exception as e:
                distances.append(0)
        
        df['distance_m'] = distances
        df['distance_km'] = df['distance_m'] / 1000
        df['cumulative_distance_km'] = cumulative_distance
        df['cumulative_fuel_l'] = cumulative_fuel
        
        # Calculate fuel per segment based on gradient
        df['gradient_class'] = df['pitch'].apply(self.classify_gradient)
        df['fuel_l'] = df.apply(lambda row: row['distance_km'] * self.fuel_rates.get(row['gradient_class'], 0.65), axis=1)
        
        df['in_maintenance'] = df.apply(lambda row: self.is_in_maintenance_area(row['lat'], row['lon']), axis=1)
        
        return df
    
    def classify_gradient(self, pitch):
        """Classify gradient based on pitch angle"""
        if pd.isna(pitch):
            return 'Flat'
        
        # Normalize pitch if needed
        if pitch > 30:
            pitch = pitch - 90
        elif pitch < -30:
            pitch = pitch + 90
            
        # Clamp to reasonable range
        pitch = max(-30, min(30, pitch))
        
        for band_name, (low, high) in self.gradient_bands.items():
            if low < pitch <= high:
                return band_name
        return 'Flat'
    
    def find_dump_and_loading_areas(self, df):
        """Use DBSCAN to find dump area and loading area based on GPS clusters"""
        valid_coords = df[(df['lat'] != 0) & (df['lon'] != 0)]
        
        if len(valid_coords) < 50:
            return None, None
        
        coords = valid_coords[['lat', 'lon']].values
        
        try:
            # Cluster the points
            eps = 0.00015  # ~15 meters
            min_samples = 20
            db = DBSCAN(eps=eps, min_samples=min_samples).fit(coords)
            valid_coords = valid_coords.copy()
            valid_coords['cluster'] = db.labels_
            
            # Get clusters (exclude noise -1)
            clustered = valid_coords[valid_coords['cluster'] != -1]
            
            if len(clustered) == 0:
                return None, None
            
            # Find the two largest clusters (dump and loading areas)
            cluster_sizes = clustered.groupby('cluster').size().sort_values(ascending=False)
            
            if len(cluster_sizes) >= 2:
                cluster1_id = cluster_sizes.index[0]
                cluster2_id = cluster_sizes.index[1]
                
                cluster1 = clustered[clustered['cluster'] == cluster1_id]
                cluster2 = clustered[clustered['cluster'] == cluster2_id]
                
                # Use altitude to determine dump (lower altitude) vs loading (higher altitude)
                alt1 = cluster1['alt'].mean() if 'alt' in cluster1.columns else 0
                alt2 = cluster2['alt'].mean() if 'alt' in cluster2.columns else 0
                
                if alt1 < alt2:
                    dump_area = (cluster1['lat'].mean(), cluster1['lon'].mean())
                    loading_area = (cluster2['lat'].mean(), cluster2['lon'].mean())
                else:
                    dump_area = (cluster2['lat'].mean(), cluster2['lon'].mean())
                    loading_area = (cluster1['lat'].mean(), cluster1['lon'].mean())
                
                print(f"    Found Dump area: ({dump_area[0]:.6f}, {dump_area[1]:.6f})", file=sys.stderr)
                print(f"    Found Loading area: ({loading_area[0]:.6f}, {loading_area[1]:.6f})", file=sys.stderr)
                
                return dump_area, loading_area
                
        except Exception as e:
            print(f"    Cluster detection error: {str(e)}", file=sys.stderr)
        
        return None, None
    
    def detect_trips_by_areas(self, df, dump_area, loading_area, radius=100):
        """Detect trips based on movement between dump and loading areas"""
        if dump_area is None or loading_area is None:
            return self.detect_trips_fallback(df)
        
        trips = []
        current_trip = 0
        in_trip = False
        at_loading = True
        
        for i in range(len(df)):
            lat, lon = df.iloc[i]['lat'], df.iloc[i]['lon']
            
            if lat == 0 or lon == 0:
                trips.append(current_trip if in_trip else 0)
                continue
            
            try:
                dist_to_dump = geodesic((lat, lon), dump_area).meters
                dist_to_loading = geodesic((lat, lon), loading_area).meters
                
                if dist_to_dump < radius:
                    current_loc = 'dump'
                elif dist_to_loading < radius:
                    current_loc = 'loading'
                else:
                    current_loc = 'enroute'
                
                if current_loc == 'enroute' and at_loading and not in_trip:
                    current_trip += 1
                    in_trip = True
                    at_loading = False
                    trips.append(current_trip)
                elif current_loc == 'dump' and not at_loading and in_trip:
                    in_trip = False
                    at_loading = False
                    trips.append(current_trip)
                elif current_loc == 'enroute' and not at_loading and not in_trip and current_loc != 'dump':
                    current_trip += 1
                    in_trip = True
                    trips.append(current_trip)
                elif current_loc == 'loading' and in_trip:
                    in_trip = False
                    at_loading = True
                    trips.append(current_trip)
                elif current_loc == 'loading':
                    at_loading = True
                    trips.append(current_trip if in_trip else 0)
                elif current_loc == 'dump':
                    at_loading = False
                    trips.append(current_trip if in_trip else 0)
                else:
                    trips.append(current_trip if in_trip else 0)
                    
            except Exception as e:
                trips.append(current_trip if in_trip else 0)
        
        if current_trip > self.max_trips_per_8h:
            trip_map = {}
            new_trip = 0
            for i, t in enumerate(trips):
                if t > 0:
                    if t not in trip_map:
                        if new_trip < self.max_trips_per_8h:
                            new_trip += 1
                            trip_map[t] = new_trip
                        else:
                            trip_map[t] = 0
                    trips[i] = trip_map[t]
        
        df['trip'] = trips
        return df
    
    def detect_trips_fallback(self, df, start_radius=50):
        """Fallback: Simple trip detection based on distance from start point"""
        valid_coords = df[(df['lat'] != 0) & (df['lon'] != 0)]
        if len(valid_coords) == 0:
            df['trip'] = 0
            return df
        
        start_lat = valid_coords.iloc[0]['lat']
        start_lon = valid_coords.iloc[0]['lon']
        start_point = (start_lat, start_lon)
        
        trips = []
        current_trip = 0
        in_trip = False
        
        for i in range(len(df)):
            cur_point = (df.iloc[i]['lat'], df.iloc[i]['lon'])
            
            if cur_point[0] == 0 and cur_point[1] == 0:
                trips.append(current_trip if in_trip else 0)
                continue
            
            try:
                dist_from_start = geodesic(cur_point, start_point).meters
                
                if dist_from_start > start_radius and not in_trip:
                    current_trip += 1
                    in_trip = True
                    trips.append(current_trip)
                elif dist_from_start < start_radius and in_trip and i > 20:
                    in_trip = False
                    trips.append(current_trip)
                else:
                    trips.append(current_trip if in_trip else 0)
            except:
                trips.append(current_trip if in_trip else 0)
        
        if current_trip > self.max_trips_per_8h:
            trip_map = {}
            new_trip = 0
            for i, t in enumerate(trips):
                if t > 0:
                    if t not in trip_map:
                        if new_trip < self.max_trips_per_8h:
                            new_trip += 1
                            trip_map[t] = new_trip
                        else:
                            trip_map[t] = 0
                    trips[i] = trip_map[t]
        
        df['trip'] = trips
        return df
    
    def detect_trips(self, df):
        dump_area, loading_area = self.find_dump_and_loading_areas(df)
        
        if dump_area and loading_area:
            print(f"    Using cluster-based trip detection (dump ↔ loading)", file=sys.stderr)
            return self.detect_trips_by_areas(df, dump_area, loading_area)
        else:
            print(f"    Using fallback trip detection (distance from start)", file=sys.stderr)
            return self.detect_trips_fallback(df)
    
    def detect_routes(self, df):
        """Detect Route A and Route B based on GPS patterns"""
        valid_coords = df[(df['lat'] != 0) & (df['lon'] != 0)]
        
        if len(valid_coords) < 100:
            return None, None
        
        coords = valid_coords[['lat', 'lon']].values
        
        try:
            eps = 0.0005
            min_samples = 50
            db = DBSCAN(eps=eps, min_samples=min_samples).fit(coords)
            valid_coords = valid_coords.copy()
            valid_coords['route_cluster'] = db.labels_
            
            route_clusters = valid_coords[valid_coords['route_cluster'] != -1]
            
            if len(route_clusters) == 0:
                return None, None
            
            cluster_sizes = route_clusters.groupby('route_cluster').size().sort_values(ascending=False)
            
            routes = {}
            for i, (cluster_id, size) in enumerate(cluster_sizes.items()):
                if i >= 2:
                    break
                
                cluster_data = route_clusters[route_clusters['route_cluster'] == cluster_id]
                
                route_distance = cluster_data['distance_km'].sum() if 'distance_km' in cluster_data.columns else 0
                route_fuel = cluster_data['fuel_l'].sum() if 'fuel_l' in cluster_data.columns else 0
                route_cost = route_fuel * self.diesel_price
                
                if 'pitch' in cluster_data.columns:
                    avg_pitch = cluster_data['pitch'].mean()
                    if avg_pitch > 8:
                        grade = "Steep Up"
                    elif avg_pitch > 3:
                        grade = "Mild Up"
                    elif avg_pitch > -3:
                        grade = "Flat"
                    elif avg_pitch > -8:
                        grade = "Mild Down"
                    else:
                        grade = "Steep Down"
                else:
                    avg_pitch = 0
                    grade = "Unknown"
                
                fuel_per_km = route_fuel / route_distance if route_distance > 0 else 0
                route_name = f"Route {chr(65 + i)}"
                
                routes[route_name] = {
                    'route_id': int(cluster_id),
                    'distance_km': round(route_distance, 2),
                    'fuel_l': round(route_fuel, 2),
                    'cost_rs': round(route_cost, 2),
                    'fuel_per_km': round(fuel_per_km, 3),
                    'avg_pitch': round(avg_pitch, 1),
                    'grade': grade,
                    'points': len(cluster_data),
                    'percentage_of_travel': round(len(cluster_data) / len(valid_coords) * 100, 1)
                }
            
            return routes.get('Route A'), routes.get('Route B')
            
        except Exception as e:
            print(f"    Route detection error: {str(e)}", file=sys.stderr)
            return None, None
    
    def analyze_device(self, device_id, df):
        """Complete analysis for a single device"""
        if len(df) < 3:
            return None
        
        try:
            required_cols = ['lat', 'lon', 'pitch', 'speed', 'alt']
            for col in required_cols:
                if col not in df.columns:
                    df[col] = 0
            
            # Clean and normalize pitch
            df['pitch'] = pd.to_numeric(df['pitch'], errors='coerce').fillna(0)
            df.loc[df['pitch'] > 30, 'pitch'] = df.loc[df['pitch'] > 30, 'pitch'] - 90
            df.loc[df['pitch'] < -30, 'pitch'] = df.loc[df['pitch'] < -30, 'pitch'] + 90
            
            # Calculate distance and fuel (using FIXED rates)
            df = self.calculate_distance_and_fuel(df)
            
            # Detect trips
            df = self.detect_trips(df)
            
            # Detect routes
            route_a, route_b = self.detect_routes(df)
            
            # Calculate maintenance time
            maintenance_time = df[df['in_maintenance'] == True]['time'].count() if 'time' in df.columns else 0
            total_time = len(df)
            maintenance_percentage = (maintenance_time / total_time * 100) if total_time > 0 else 0
            
            # Filter to valid GPS points
            valid_gps = df[(df['lat'] != 0) & (df['lon'] != 0) & (df['distance_km'] > 0)]
            
            if len(valid_gps) == 0:
                print(f"  Device {device_id}: No valid GPS data", file=sys.stderr)
                return None
            
            # Basic metrics
            total_distance = valid_gps['distance_km'].sum()
            total_fuel = valid_gps['fuel_l'].sum()
            total_cost = total_fuel * self.diesel_price
            
            if total_fuel == 0:
                print(f"  Device {device_id}: No fuel consumption calculated", file=sys.stderr)
                return None
            
            print(f"  Device {device_id} - Distance: {total_distance:.2f}km, Fuel: {total_fuel:.2f}L, Cost: ₹{total_cost:.2f}", file=sys.stderr)
            
            # Trip analysis
            unique_trips = df[df['trip'] > 0]['trip'].nunique()
            if unique_trips > self.max_trips_per_8h:
                unique_trips = self.max_trips_per_8h
            
            trip_details = []
            total_trip_distance = 0
            total_trip_fuel = 0
            
            for trip_num in sorted(df[df['trip'] > 0]['trip'].unique()):
                if trip_num > self.max_trips_per_8h:
                    continue
                    
                trip_data = df[df['trip'] == trip_num]
                
                if len(trip_data) < 5:
                    continue
                
                trip_distance = trip_data['distance_km'].sum()
                if trip_distance == 0:
                    continue
                
                trip_fuel = trip_data['fuel_l'].sum()
                trip_cost = trip_fuel * self.diesel_price
                
                total_trip_distance += trip_distance
                total_trip_fuel += trip_fuel
                
                trip_duration = None
                if 'time' in trip_data.columns and len(trip_data) > 1:
                    trip_duration = trip_data.iloc[-1]['time'] - trip_data.iloc[0]['time']
                
                avg_speed = trip_data['speed'].mean() if 'speed' in trip_data.columns else 0
                if avg_speed > self.max_realistic_speed:
                    avg_speed = self.max_realistic_speed
                
                trip_details.append({
                    'trip_number': int(trip_num),
                    'start_time': str(trip_data.iloc[0]['time']) if 'time' in trip_data.columns else 'N/A',
                    'end_time': str(trip_data.iloc[-1]['time']) if 'time' in trip_data.columns else 'N/A',
                    'duration_hours': round(trip_duration.total_seconds() / 3600, 2) if trip_duration else 0,
                    'distance_km': round(float(trip_distance), 2),
                    'fuel_l': round(float(trip_fuel), 2),
                    'cost_rs': round(float(trip_cost), 2),
                    'avg_speed': round(float(avg_speed), 1)
                })
            
            # Average speed
            moving = df[(df['speed'] > 2) & (df['lat'] != 0)]
            avg_speed = float(moving['speed'].mean()) if len(moving) > 0 else 0
            if avg_speed > self.max_realistic_speed:
                avg_speed = self.max_realistic_speed
            
            # Fuel efficiency
            fuel_per_km = total_fuel / total_distance if total_distance > 0 else 0
            
            # Gradient analysis - Simplified to Steep Up, Steep Down, Flat only
            gradient_stats = {}
            
            steep_up_mask = df['gradient_class'] == 'Steep Up'
            steep_up_data = df[steep_up_mask]
            steep_up_distance = steep_up_data['distance_km'].sum()
            steep_up_fuel = steep_up_data['fuel_l'].sum()
            
            steep_down_mask = df['gradient_class'] == 'Steep Down'
            steep_down_data = df[steep_down_mask]
            steep_down_distance = steep_down_data['distance_km'].sum()
            steep_down_fuel = steep_down_data['fuel_l'].sum()
            
            flat_mask = df['gradient_class'] == 'Flat'
            flat_data = df[flat_mask]
            flat_distance = flat_data['distance_km'].sum()
            flat_fuel = flat_data['fuel_l'].sum()
            
            # Also include Mild Up/Down in Flat for simplicity
            mild_up_mask = df['gradient_class'] == 'Mild Up'
            mild_up_data = df[mild_up_mask]
            flat_distance += mild_up_data['distance_km'].sum()
            flat_fuel += mild_up_data['fuel_l'].sum()
            
            mild_down_mask = df['gradient_class'] == 'Mild Down'
            mild_down_data = df[mild_down_mask]
            flat_distance += mild_down_data['distance_km'].sum()
            flat_fuel += mild_down_data['fuel_l'].sum()
            
            gradient_stats['Steep Up'] = {
                'distance_km': round(steep_up_distance, 2),
                'fuel_l': round(steep_up_fuel, 2),
                'cost_rs': round(steep_up_fuel * self.diesel_price, 2),
                'percentage': round(steep_up_distance / total_distance * 100, 1) if total_distance > 0 else 0
            }
            
            gradient_stats['Steep Down'] = {
                'distance_km': round(steep_down_distance, 2),
                'fuel_l': round(steep_down_fuel, 2),
                'cost_rs': round(steep_down_fuel * self.diesel_price, 2),
                'percentage': round(steep_down_distance / total_distance * 100, 1) if total_distance > 0 else 0
            }
            
            gradient_stats['Flat'] = {
                'distance_km': round(flat_distance, 2),
                'fuel_l': round(flat_fuel, 2),
                'cost_rs': round(flat_fuel * self.diesel_price, 2),
                'percentage': round(flat_distance / total_distance * 100, 1) if total_distance > 0 else 0
            }
            
            # Idle analysis
            idle_points = len(df[(df['speed'] <= 2) & (df['lat'] != 0)]) if 'speed' in df.columns else 0
            total_points = len(df[(df['lat'] != 0)])
            idle_ratio = idle_points / total_points if total_points > 0 else 0
            
            if idle_ratio > 0.4:
                idle_pattern = "High Idle Time"
            elif idle_ratio > 0.2:
                idle_pattern = "Medium Idle Time"
            else:
                idle_pattern = "Low Idle Time"
            
            # Route type based on steep gradients
            steep_up_pct = gradient_stats.get('Steep Up', {}).get('percentage', 0)
            
            if steep_up_pct > 30:
                route_type = "Steep Route (High Grade)"
                route_chars = "High fuel consumption, low speed, frequent steep climbs"
            elif steep_up_pct > 15:
                route_type = "Mixed Terrain"
                route_chars = "Moderate fuel consumption, balanced speed"
            else:
                route_type = "Gentle Route"
                route_chars = "Optimal fuel efficiency, good speed"
            
            # Performance rating based on ACTUAL fuel efficiency
            if fuel_per_km < 0.4:
                fuel_efficiency_rating = "Excellent"
            elif fuel_per_km < 0.65:
                fuel_efficiency_rating = "Good"
            elif fuel_per_km < 0.85:
                fuel_efficiency_rating = "Average"
            else:
                fuel_efficiency_rating = "Needs Improvement"
            
            return {
                'device_id': str(device_id),
                'total_distance': round(total_distance, 1),
                'total_fuel': round(total_fuel, 1),
                'total_cost': round(total_cost, 2),
                'fuel_per_km': round(fuel_per_km, 2),
                'trips': unique_trips,
                'total_trip_distance': round(total_trip_distance, 1),
                'total_trip_fuel': round(total_trip_fuel, 1),
                'avg_speed': round(avg_speed, 1),
                'fuel_efficiency_rating': fuel_efficiency_rating,
                'gradient_stats': gradient_stats,
                'idle_pattern': idle_pattern,
                'idle_ratio': round(idle_ratio * 100, 1),
                'route_type': route_type,
                'route_chars': route_chars,
                'steep_up_percentage': round(steep_up_pct, 1),
                'trip_details': trip_details[:20],
                'route_a': route_a,
                'route_b': route_b,
                'maintenance_time_pct': round(maintenance_percentage, 1),
                'in_maintenance': maintenance_percentage > 0
            }
            
        except Exception as e:
            print(f"Error analyzing device {device_id}: {str(e)}", file=sys.stderr)
            import traceback
            traceback.print_exc(file=sys.stderr)
            return None


class ExcelReportGenerator:
    """Generate comprehensive Excel report"""
    
    def __init__(self):
        self.wb = Workbook()
        self.setup_styles()
        
        self.device_mapping = {
            'D3': '135',
            'D8': '133',
            'D9': '137',
            'D7': '07',
            'D10': '08',
            'D12': '134'
        }
        
        self.excavator_ids = ['07', 'D7', '7']
    
    def map_device_id(self, device_id):
        return self.device_mapping.get(str(device_id).upper(), device_id)
    
    def is_excavator(self, device_id):
        device_str = str(device_id).upper()
        for exc_id in self.excavator_ids:
            if exc_id.upper() in device_str:
                return True
        return False
    
    def setup_styles(self):
        self.header_fill = PatternFill(start_color='2C3E50', end_color='2C3E50', fill_type='solid')
        self.header_font = Font(color='FFFFFF', bold=True, size=11)
        self.centered = Alignment(horizontal='center', vertical='center')
        self.right_aligned = Alignment(horizontal='right', vertical='center')
        self.border = Border(
            left=Side(style='thin', color='BDC3C7'),
            right=Side(style='thin', color='BDC3C7'),
            top=Side(style='thin', color='BDC3C7'),
            bottom=Side(style='thin', color='BDC3C7')
        )
        self.good_fill = PatternFill(start_color='2ECC71', end_color='2ECC71', fill_type='solid')
        self.bad_fill = PatternFill(start_color='E74C3C', end_color='E74C3C', fill_type='solid')
        self.warning_fill = PatternFill(start_color='F39C12', end_color='F39C12', fill_type='solid')
        self.maintenance_fill = PatternFill(start_color='3498DB', end_color='3498DB', fill_type='solid')
    
    def create_executive_summary(self, results):
        ws = self.wb.create_sheet("Executive Summary", 0)
        
        title_cell = ws.cell(row=1, column=1, value="MINING HAULER PERFORMANCE DASHBOARD")
        title_cell.font = Font(size=16, bold=True, color='2C3E50')
        title_cell.alignment = self.centered
        ws.merge_cells('A1:F1')
        
        date_cell = ws.cell(row=2, column=1, value=f"Report Generated: {datetime.now().strftime('%d %B %Y %H:%M')}")
        date_cell.font = Font(size=10, italic=True)
        ws.merge_cells('A2:F2')
        
        maint_note = ws.cell(row=3, column=1, value="📍 TMC Maintenance Area: Lat 20.4051928, Lon 81.0662203")
        maint_note.font = Font(size=9, color='3498DB')
        ws.merge_cells('A3:F3')
        
        hauler_results = {k: v for k, v in results.items() if not self.is_excavator(k)}
        
        total_fuel = sum(r['total_fuel'] for r in hauler_results.values())
        total_cost = sum(r['total_cost'] for r in hauler_results.values())
        total_distance = sum(r['total_distance'] for r in hauler_results.values())
        total_trips = sum(r['trips'] for r in hauler_results.values())
        avg_fuel_per_km = total_fuel / total_distance if total_distance > 0 else 0
        
        metrics = [
            ['Total Fuel Consumed', f"{total_fuel:,.1f} Liters", f"₹{total_cost:,.2f}"],
            ['Total Distance Traveled', f"{total_distance:,.1f} km", ''],
            ['Total Trips Completed', f"{total_trips}", ''],
            ['Average Fuel Efficiency', f"{avg_fuel_per_km:.2f} L/km", ''],
            ['Number of Haulers Analyzed', f"{len(hauler_results)}", '']
        ]
        
        row = 5
        for metric in metrics:
            ws.cell(row=row, column=1, value=metric[0]).font = Font(bold=True)
            ws.cell(row=row, column=2, value=metric[1])
            if len(metric) > 2:
                ws.cell(row=row, column=3, value=metric[2])
            row += 1
        
        for col in range(1, 4):
            ws.column_dimensions[get_column_letter(col)].width = 25
        
        return ws
    
    def create_hauler_performance(self, results):
        ws = self.wb.create_sheet("Hauler Performance")
        
        headers = ['Hauler ID', 'Distance (km)', 'Fuel (L)', 'Cost (₹)', 'Fuel/km (L/km)', 
                   'Trips', 'Avg Speed (km/h)', 'Efficiency Rating', 'Route Type', 'Maintenance %']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.centered
            cell.border = self.border
        
        row = 2
        for device_id, data in sorted(results.items(), key=lambda x: x[1]['total_fuel'], reverse=True):
            if self.is_excavator(device_id):
                continue
            
            ws.cell(row=row, column=1, value=self.map_device_id(device_id))
            ws.cell(row=row, column=2, value=data['total_distance'])
            ws.cell(row=row, column=3, value=data['total_fuel'])
            ws.cell(row=row, column=4, value=data['total_cost'])
            ws.cell(row=row, column=5, value=data['fuel_per_km'])
            ws.cell(row=row, column=6, value=data['trips'])
            ws.cell(row=row, column=7, value=data['avg_speed'])
            ws.cell(row=row, column=8, value=data['fuel_efficiency_rating'])
            ws.cell(row=row, column=9, value=data['route_type'])
            ws.cell(row=row, column=10, value=data.get('maintenance_time_pct', 0))
            
            if data['fuel_efficiency_rating'] == 'Excellent':
                ws.cell(row=row, column=8).fill = self.good_fill
            elif data['fuel_efficiency_rating'] == 'Needs Improvement':
                ws.cell(row=row, column=8).fill = self.bad_fill
            
            if data.get('maintenance_time_pct', 0) > 0:
                ws.cell(row=row, column=10).fill = self.maintenance_fill
            
            for col in range(1, 11):
                ws.cell(row=row, column=col).border = self.border
                if col not in [1, 8, 9]:
                    ws.cell(row=row, column=col).alignment = self.right_aligned
            
            row += 1
        
        for col in range(1, 11):
            ws.column_dimensions[get_column_letter(col)].width = 14
        
        return ws
    
    def create_excavator_analysis(self, results):
        ws = self.wb.create_sheet("Excavator Analysis")
        
        headers = ['Equipment ID', 'Type', 'Total Distance (km)', 'Fuel (L)', 'Cost (₹)', 
                   'Fuel/km (L/km)', 'Avg Speed (km/h)', 'Idle %', 'Maintenance %', 'Operating Status']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.centered
            cell.border = self.border
        
        row = 2
        excavator_found = False
        for device_id, data in results.items():
            if self.is_excavator(device_id):
                excavator_found = True
                operating_status = "Stationary Operation" if data['total_distance'] < 5 else "Moving Operation"
                
                ws.cell(row=row, column=1, value=self.map_device_id(device_id))
                ws.cell(row=row, column=2, value="Excavator")
                ws.cell(row=row, column=3, value=data['total_distance'])
                ws.cell(row=row, column=4, value=data['total_fuel'])
                ws.cell(row=row, column=5, value=data['total_cost'])
                ws.cell(row=row, column=6, value=data['fuel_per_km'])
                ws.cell(row=row, column=7, value=data['avg_speed'])
                ws.cell(row=row, column=8, value=data['idle_ratio'])
                ws.cell(row=row, column=9, value=data.get('maintenance_time_pct', 0))
                ws.cell(row=row, column=10, value=operating_status)
                
                for col in range(1, 11):
                    ws.cell(row=row, column=col).border = self.border
                    if col > 1:
                        ws.cell(row=row, column=col).alignment = self.right_aligned
                row += 1
        
        if not excavator_found:
            ws.cell(row=2, column=1, value="No excavator data available")
        
        for col in range(1, 11):
            ws.column_dimensions[get_column_letter(col)].width = 16
        
        return ws
    
    def create_route_analysis(self, results):
        ws = self.wb.create_sheet("Route Analysis")
        
        ws.merge_cells('A1:K1')
        title_cell = ws.cell(row=1, column=1, value="ROUTE COMPARISON: ROUTE A vs ROUTE B")
        title_cell.font = Font(size=14, bold=True, color='2C3E50')
        title_cell.alignment = self.centered
        
        ws.merge_cells('A2:K2')
        ws.cell(row=2, column=1, value="Comparing fuel efficiency, distance, and gradient between routes")
        ws.cell(row=2, column=1).alignment = self.centered
        
        headers = ['Hauler ID', 'Route', 'Distance (km)', 'Fuel (L)', 'Cost (₹)', 
                   'Fuel/km (L/km)', 'Avg Pitch (°)', 'Grade', '% of Travel', 'Recommendation']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.centered
            cell.border = self.border
        
        row = 5
        route_comparisons = []
        
        for device_id, data in results.items():
            if self.is_excavator(device_id):
                continue
            
            route_a = data.get('route_a')
            route_b = data.get('route_b')
            
            if route_a:
                recommendation = ""
                if route_b:
                    if route_a['fuel_per_km'] < route_b['fuel_per_km']:
                        recommendation = "Route A is more fuel efficient"
                    elif route_b['fuel_per_km'] < route_a['fuel_per_km']:
                        recommendation = "Route B is more fuel efficient"
                    else:
                        recommendation = "Both routes have similar efficiency"
                    
                    route_comparisons.append({
                        'device': device_id,
                        'route_a_fuel': route_a['fuel_per_km'],
                        'route_b_fuel': route_b['fuel_per_km'],
                        'better_route': 'A' if route_a['fuel_per_km'] < route_b['fuel_per_km'] else 'B'
                    })
                else:
                    recommendation = "Only Route A detected"
                
                ws.cell(row=row, column=1, value=self.map_device_id(device_id))
                ws.cell(row=row, column=2, value="Route A")
                ws.cell(row=row, column=3, value=route_a['distance_km'])
                ws.cell(row=row, column=4, value=route_a['fuel_l'])
                ws.cell(row=row, column=5, value=route_a['cost_rs'])
                ws.cell(row=row, column=6, value=route_a['fuel_per_km'])
                ws.cell(row=row, column=7, value=route_a['avg_pitch'])
                ws.cell(row=row, column=8, value=route_a['grade'])
                ws.cell(row=row, column=9, value=route_a['percentage_of_travel'])
                ws.cell(row=row, column=10, value=recommendation)
                
                if route_a['fuel_per_km'] < 0.6:
                    ws.cell(row=row, column=6).fill = self.good_fill
                elif route_a['fuel_per_km'] > 0.9:
                    ws.cell(row=row, column=6).fill = self.bad_fill
                
                for col in range(1, 11):
                    ws.cell(row=row, column=col).border = self.border
                    if col not in [1, 2, 8, 10]:
                        ws.cell(row=row, column=col).alignment = self.right_aligned
                row += 1
            
            if route_b:
                ws.cell(row=row, column=1, value=self.map_device_id(device_id))
                ws.cell(row=row, column=2, value="Route B")
                ws.cell(row=row, column=3, value=route_b['distance_km'])
                ws.cell(row=row, column=4, value=route_b['fuel_l'])
                ws.cell(row=row, column=5, value=route_b['cost_rs'])
                ws.cell(row=row, column=6, value=route_b['fuel_per_km'])
                ws.cell(row=row, column=7, value=route_b['avg_pitch'])
                ws.cell(row=row, column=8, value=route_b['grade'])
                ws.cell(row=row, column=9, value=route_b['percentage_of_travel'])
                ws.cell(row=row, column=10, value="")
                
                if route_b['fuel_per_km'] < 0.6:
                    ws.cell(row=row, column=6).fill = self.good_fill
                elif route_b['fuel_per_km'] > 0.9:
                    ws.cell(row=row, column=6).fill = self.bad_fill
                
                for col in range(1, 11):
                    ws.cell(row=row, column=col).border = self.border
                    if col not in [1, 2, 8, 10]:
                        ws.cell(row=row, column=col).alignment = self.right_aligned
                row += 1
            
            if route_a or route_b:
                row += 1
        
        if route_comparisons:
            row += 2
            ws.merge_cells(f'A{row}:K{row}')
            ws.cell(row=row, column=1, value="SUMMARY: ROUTE EFFICIENCY COMPARISON")
            ws.cell(row=row, column=1).font = Font(bold=True, size=12)
            ws.cell(row=row, column=1).alignment = self.centered
            
            row += 1
            comp_headers = ['Hauler ID', 'Route A (L/km)', 'Route B (L/km)', 'Better Route', 'Fuel Saving (L/km)']
            for col, header in enumerate(comp_headers, 1):
                cell = ws.cell(row=row, column=col, value=header)
                cell.font = self.header_font
                cell.fill = self.header_fill
                cell.alignment = self.centered
                cell.border = self.border
            
            row += 1
            for comp in route_comparisons:
                saving = abs(comp['route_a_fuel'] - comp['route_b_fuel'])
                ws.cell(row=row, column=1, value=self.map_device_id(comp['device']))
                ws.cell(row=row, column=2, value=round(comp['route_a_fuel'], 3))
                ws.cell(row=row, column=3, value=round(comp['route_b_fuel'], 3))
                ws.cell(row=row, column=4, value=f"Route {comp['better_route']}")
                ws.cell(row=row, column=5, value=round(saving, 3))
                
                for col in range(1, 6):
                    ws.cell(row=row, column=col).border = self.border
                    if col > 1:
                        ws.cell(row=row, column=col).alignment = self.right_aligned
                row += 1
        
        for col in range(1, 11):
            ws.column_dimensions[get_column_letter(col)].width = 14
        ws.column_dimensions[get_column_letter(10)].width = 30
        
        return ws
    
    def create_gradient_analysis(self, results):
        """SIMPLIFIED Terrain Analysis - Only Steep Up, Steep Down, Flat"""
        ws = self.wb.create_sheet("Terrain Analysis")
        
        headers = ['Hauler ID', 'Steep Up (km)', 'Steep Up Fuel (L)', 
                   'Steep Down (km)', 'Steep Down Fuel (L)',
                   'Flat (km)', 'Flat Fuel (L)', 'Steep Grade %', 'Efficiency on Steep']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.centered
            cell.border = self.border
        
        row = 2
        for device_id, data in results.items():
            if self.is_excavator(device_id):
                continue
            
            grad_stats = data['gradient_stats']
            steep_up_dist = grad_stats.get('Steep Up', {}).get('distance_km', 0)
            steep_up_fuel = grad_stats.get('Steep Up', {}).get('fuel_l', 0)
            
            # Calculate efficiency on steep terrain (L/km)
            steep_efficiency = steep_up_fuel / steep_up_dist if steep_up_dist > 0 else 0
            
            ws.cell(row=row, column=1, value=self.map_device_id(device_id))
            ws.cell(row=row, column=2, value=steep_up_dist)
            ws.cell(row=row, column=3, value=steep_up_fuel)
            ws.cell(row=row, column=4, value=grad_stats.get('Steep Down', {}).get('distance_km', 0))
            ws.cell(row=row, column=5, value=grad_stats.get('Steep Down', {}).get('fuel_l', 0))
            ws.cell(row=row, column=6, value=grad_stats.get('Flat', {}).get('distance_km', 0))
            ws.cell(row=row, column=7, value=grad_stats.get('Flat', {}).get('fuel_l', 0))
            ws.cell(row=row, column=8, value=grad_stats.get('Steep Up', {}).get('percentage', 0))
            ws.cell(row=row, column=9, value=round(steep_efficiency, 2))
            
            # Color coding
            if steep_efficiency > 1.0:
                ws.cell(row=row, column=9).fill = self.bad_fill
            elif steep_efficiency > 0.7:
                ws.cell(row=row, column=9).fill = self.warning_fill
            else:
                ws.cell(row=row, column=9).fill = self.good_fill
            
            for col in range(1, 10):
                ws.cell(row=row, column=col).border = self.border
                if col > 1:
                    ws.cell(row=row, column=col).alignment = self.right_aligned
            row += 1
        
        for col in range(1, 10):
            ws.column_dimensions[get_column_letter(col)].width = 14
        
        # Add explanation
        row += 2
        ws.cell(row=row, column=1, value="📊 Fuel Efficiency Explanation:")
        ws.cell(row=row, column=1).font = Font(bold=True)
        row += 1
        ws.cell(row=row, column=1, value="• Steep Up > 1.0 L/km = Needs Improvement")
        ws.cell(row=row, column=1).fill = self.bad_fill
        row += 1
        ws.cell(row=row, column=1, value="• Steep Up 0.7-1.0 L/km = Average")
        ws.cell(row=row, column=1).fill = self.warning_fill
        row += 1
        ws.cell(row=row, column=1, value="• Steep Up < 0.7 L/km = Good/Excellent")
        ws.cell(row=row, column=1).fill = self.good_fill
        
        return ws
    
    def create_trip_analysis(self, results):
        ws = self.wb.create_sheet("Trip Summary")
        
        headers = ['Hauler ID', 'Total Trips', 'Total Distance (km)', 'Total Fuel (L)',
                   'Avg Trip Distance (km)', 'Avg Trip Fuel (L)', 'Avg Speed (km/h)']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.centered
            cell.border = self.border
        
        row = 2
        for device_id, data in results.items():
            if self.is_excavator(device_id):
                continue
            
            total_trips = data.get('trips', 0)
            total_distance = data.get('total_trip_distance', 0)
            total_fuel = data.get('total_trip_fuel', 0)
            
            avg_trip_distance = total_distance / total_trips if total_trips > 0 else 0
            avg_trip_fuel = total_fuel / total_trips if total_trips > 0 else 0
            
            ws.cell(row=row, column=1, value=self.map_device_id(device_id))
            ws.cell(row=row, column=2, value=total_trips)
            ws.cell(row=row, column=3, value=round(total_distance, 2))
            ws.cell(row=row, column=4, value=round(total_fuel, 2))
            ws.cell(row=row, column=5, value=round(avg_trip_distance, 2))
            ws.cell(row=row, column=6, value=round(avg_trip_fuel, 2))
            ws.cell(row=row, column=7, value=data.get('avg_speed', 0))
            
            for col in range(1, 8):
                ws.cell(row=row, column=col).border = self.border
                if col > 1:
                    ws.cell(row=row, column=col).alignment = self.right_aligned
            row += 1
        
        return ws
    
    def create_operational_insights(self, results):
        ws = self.wb.create_sheet("Operational Insights")
        
        headers = ['Hauler ID', 'Idle Pattern', 'Idle %', 'Route Type', 'Avg Speed (km/h)', 
                   'Fuel Efficiency', 'Maintenance %', 'Recommendation']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.centered
            cell.border = self.border
        
        row = 2
        for device_id, data in results.items():
            if self.is_excavator(device_id):
                continue
            
            recommendation = []
            if data['fuel_per_km'] > 0.85:
                recommendation.append(f"High fuel consumption ({data['fuel_per_km']} L/km) - check steep routes")
            if data['idle_ratio'] > 30:
                recommendation.append("Excessive idle time")
            if data['steep_up_percentage'] > 30:
                recommendation.append(f"Steep terrain {data['steep_up_percentage']}% of journey")
            if data['avg_speed'] < 5:
                recommendation.append("Low average speed - check vehicle condition")
            if data.get('maintenance_time_pct', 0) > 10:
                recommendation.append(f"In maintenance {data.get('maintenance_time_pct', 0)}% of time")
            
            if data.get('route_a') and data.get('route_b'):
                if data['route_a']['fuel_per_km'] < data['route_b']['fuel_per_km']:
                    recommendation.append("Use Route A for better fuel efficiency")
                else:
                    recommendation.append("Use Route B for better fuel efficiency")
            
            rec_text = "; ".join(recommendation) if recommendation else "Operating within normal parameters"
            
            ws.cell(row=row, column=1, value=self.map_device_id(device_id))
            ws.cell(row=row, column=2, value=data['idle_pattern'])
            ws.cell(row=row, column=3, value=data['idle_ratio'])
            ws.cell(row=row, column=4, value=data['route_type'])
            ws.cell(row=row, column=5, value=data['avg_speed'])
            ws.cell(row=row, column=6, value=data['fuel_efficiency_rating'])
            ws.cell(row=row, column=7, value=data.get('maintenance_time_pct', 0))
            ws.cell(row=row, column=8, value=rec_text)
            
            if data.get('maintenance_time_pct', 0) > 10:
                ws.cell(row=row, column=7).fill = self.maintenance_fill
            
            for col in range(1, 9):
                ws.cell(row=row, column=col).border = self.border
            row += 1
        
        for col in range(1, 9):
            ws.column_dimensions[get_column_letter(col)].width = 16
        ws.column_dimensions[get_column_letter(8)].width = 45
        
        return ws
    
    def generate_report(self, results):
        if 'Sheet' in self.wb.sheetnames:
            std = self.wb['Sheet']
            self.wb.remove(std)
        
        self.create_executive_summary(results)
        self.create_hauler_performance(results)
        self.create_excavator_analysis(results)
        self.create_route_analysis(results)
        self.create_gradient_analysis(results)
        self.create_trip_analysis(results)
        self.create_operational_insights(results)
        
        return self.wb


def main():
    try:
        input_data = sys.stdin.read()
        if not input_data:
            print(json.dumps({'status': 'error', 'error': 'No input data'}))
            return
        
        params = json.loads(input_data)
        df = pd.DataFrame(params.get('data', []))
        
        if df.empty:
            print(json.dumps({
                'status': 'success',
                'report': '',
                'filename': 'No_Data.xlsx',
                'devices_discovered': [],
                'record_count': 0
            }))
            return
        
        print(f"Raw data columns: {df.columns.tolist()}", file=sys.stderr)
        print(f"Total records: {len(df)}", file=sys.stderr)
        
        essential_cols = ['lat', 'lon', 'pitch', 'speed']
        missing_cols = [col for col in essential_cols if col not in df.columns]
        if missing_cols:
            print(f"Warning: Missing columns: {missing_cols}", file=sys.stderr)
            for col in missing_cols:
                df[col] = 0
        
        if 'device_id' not in df.columns:
            df['device_id'] = 'UNKNOWN'
        
        if 'time' in df.columns:
            df['time'] = pd.to_datetime(df['time'])
            df = df.sort_values('time').reset_index(drop=True)
        
        df['lat'] = pd.to_numeric(df['lat'], errors='coerce').fillna(0)
        df['lon'] = pd.to_numeric(df['lon'], errors='coerce').fillna(0)
        df['pitch'] = pd.to_numeric(df['pitch'], errors='coerce').fillna(0)
        df['speed'] = pd.to_numeric(df['speed'], errors='coerce').fillna(0)
        
        print(f"\n{'='*50}", file=sys.stderr)
        print(f"Processing {len(df)} records", file=sys.stderr)
        print(f"Unique devices: {df['device_id'].unique().tolist()}", file=sys.stderr)
        
        analyzer = SimpleMiningAnalytics()
        devices = df['device_id'].unique()
        results = {}
        
        for device_id in devices:
            if pd.isna(device_id):
                continue
            
            device_df = df[df['device_id'] == device_id].copy()
            
            if len(device_df) < 3:
                print(f"Device {device_id}: insufficient data ({len(device_df)} records)", file=sys.stderr)
                continue
            
            print(f"\n{'='*50}", file=sys.stderr)
            print(f"Analyzing device {device_id} ({len(device_df)} records)", file=sys.stderr)
            
            result = analyzer.analyze_device(device_id, device_df)
            if result:
                results[str(device_id)] = result
                print(f"  ✓ Device {device_id}: {result['trips']} trips, {result['total_distance']} km, {result['total_fuel']} L fuel", file=sys.stderr)
                if result.get('maintenance_time_pct', 0) > 0:
                    print(f"     📍 Maintenance area: {result['maintenance_time_pct']}% of time", file=sys.stderr)
        
        if not results:
            print("\n❌ No valid results generated", file=sys.stderr)
            print(json.dumps({
                'status': 'success',
                'report': '',
                'filename': 'No_Valid_Data.xlsx',
                'devices_discovered': [],
                'record_count': len(df),
                'warning': 'No devices with valid data'
            }))
            return
        
        print(f"\n{'='*50}", file=sys.stderr)
        print("Generating Excel report...", file=sys.stderr)
        generator = ExcelReportGenerator()
        wb = generator.generate_report(results)
        
        excel_bytes = BytesIO()
        wb.save(excel_bytes)
        excel_bytes.seek(0)
        
        date_str = datetime.now().strftime("%d%b%y").upper()
        filename = f"Mining_Analytics_{date_str}.xlsx"
        
        hauler_results = {k: v for k, v in results.items() if not generator.is_excavator(k)}
        
        output = {
            'report': base64.b64encode(excel_bytes.read()).decode('utf-8'),
            'filename': filename,
            'status': 'success',
            'devices_discovered': list(results.keys()),
            'record_count': len(df),
            'summary': {
                'total_fuel': sum(r['total_fuel'] for r in hauler_results.values()),
                'total_cost': sum(r['total_cost'] for r in hauler_results.values()),
                'total_distance': sum(r['total_distance'] for r in hauler_results.values()),
                'total_trips': sum(r['trips'] for r in hauler_results.values())
            }
        }
        
        print(f"\n✅ Report generated successfully!", file=sys.stderr)
        print(f"   Total Fuel: {output['summary']['total_fuel']:,.2f} L", file=sys.stderr)
        print(f"   Total Cost: ₹{output['summary']['total_cost']:,.2f}", file=sys.stderr)
        
        print(json.dumps(output))
        
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        print(f"ERROR: {str(e)}", file=sys.stderr)
        print(error_trace, file=sys.stderr)
        print(json.dumps({'status': 'error', 'error': str(e), 'traceback': error_trace}))

if __name__ == '__main__':
    main()