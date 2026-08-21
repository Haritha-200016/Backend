# routes/analysis.py - COMPLETE CORRECTED CODE WITH REALISTIC TRIPS PER HOUR
'''
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
    def __init__(self):
        self.diesel_price = 94.5
        self.ton_per_trip = 35
        
        self.fuel_rates = {
            'Steep Up': 1.05,
            'Mild Up': 0.85,
            'Flat': 0.65,
            'Mild Down': 0.45,
            'Steep Down': 0.25
        }
        
        self.max_realistic_speed = 35
        self.max_trips_per_shift = 13
        
        self.gradient_bands = {
            'Steep Down': (-float('inf'), -8),
            'Mild Down': (-8, -3),
            'Flat': (-3, 3),
            'Mild Up': (3, 8),
            'Steep Up': (8, float('inf'))
        }
        
        self.maintenance_lat = 20.4051928
        self.maintenance_lon = 81.0662203
        self.maintenance_radius = 100
        
        self.working_area_lat = 20.4088
        self.working_area_lon = 81.0655
        self.working_area_radius = 150
        
        self.min_trip_distance_km = 0.5
        
        self.hauler_ids = ['D3', 'D10', 'D11', 'D12', '137', '133', '134', '135']
        self.excavator_ids = ['D7', '07', '7', '43']
        self.bulldozer_ids = ['D8', '8']
        
        self.device_mapping = {
            'D3': '137', 'D10': '133', 'D11': '134', 'D7': '43', 'D12': '135'
        }
    
    def map_device_id(self, device_id):
        return self.device_mapping.get(str(device_id).upper(), device_id)
    
    def is_hauler(self, device_id):
        device_str = str(device_id).upper()
        for hid in self.hauler_ids:
            if hid.upper() in device_str:
                return True
        return False
    
    def is_excavator(self, device_id):
        device_str = str(device_id).upper()
        for eid in self.excavator_ids:
            if eid.upper() in device_str:
                return True
        return False
    
    def is_bulldozer(self, device_id):
        device_str = str(device_id).upper()
        for bid in self.bulldozer_ids:
            if bid.upper() in device_str:
                return True
        return False
    
    def is_in_maintenance_area(self, lat, lon):
        if lat == 0 or lon == 0:
            return False
        try:
            dist = geodesic((lat, lon), (self.maintenance_lat, self.maintenance_lon)).meters
            return dist <= self.maintenance_radius
        except:
            return False
    
    def is_in_working_area(self, lat, lon):
        if lat == 0 or lon == 0:
            return False
        try:
            dist = geodesic((lat, lon), (self.working_area_lat, self.working_area_lon)).meters
            return dist <= self.working_area_radius
        except:
            return False
    
    def classify_gradient(self, pitch):
        if pd.isna(pitch):
            return 'Flat'
        if pitch > 30:
            pitch = pitch - 90
        elif pitch < -30:
            pitch = pitch + 90
        pitch = max(-30, min(30, pitch))
        for band_name, (low, high) in self.gradient_bands.items():
            if low < pitch <= high:
                return band_name
        return 'Flat'
    
    def calculate_distance_and_fuel(self, df):
        distances = []
        cumulative_distance = 0
        cumulative_fuel = 0
        
        for i in range(len(df)):
            if i == 0:
                distances.append(0)
                continue
                
            p1 = (df.iloc[i-1]['lat'], df.iloc[i-1]['lon'])
            p2 = (df.iloc[i]['lat'], df.iloc[i]['lon'])
            
            if p1[0] == 0 or p1[1] == 0 or p2[0] == 0 or p2[1] == 0:
                distances.append(0)
                continue
            
            try:
                dist_m = geodesic(p1, p2).meters
                if dist_m > 5000:
                    distances.append(0)
                    continue
                if dist_m < 1:
                    distances.append(0)
                    continue
                    
                distances.append(dist_m)
                dist_km = dist_m / 1000
                cumulative_distance += dist_km
                
                pitch = df.iloc[i]['pitch']
                gradient_class = self.classify_gradient(pitch)
                fuel_rate = self.fuel_rates.get(gradient_class, 0.65)
                segment_fuel = dist_km * fuel_rate
                cumulative_fuel += segment_fuel
                
            except:
                distances.append(0)
        
        df['distance_m'] = distances
        df['distance_km'] = df['distance_m'] / 1000
        df['cumulative_distance_km'] = cumulative_distance
        df['cumulative_fuel_l'] = cumulative_fuel
        df['gradient_class'] = df['pitch'].apply(self.classify_gradient)
        df['fuel_l'] = df.apply(lambda row: row['distance_km'] * self.fuel_rates.get(row['gradient_class'], 0.65), axis=1)
        df['in_maintenance'] = df.apply(lambda row: self.is_in_maintenance_area(row['lat'], row['lon']), axis=1)
        df['in_working_area'] = df.apply(lambda row: self.is_in_working_area(row['lat'], row['lon']), axis=1)
        
        return df
    
    def detect_real_trips(self, df):
        trips = []
        current_trip = 0
        in_trip = False
        trip_distance = 0
        left_working_area = False
        
        for i in range(len(df)):
            lat, lon = df.iloc[i]['lat'], df.iloc[i]['lon']
            in_maintenance = df.iloc[i]['in_maintenance']
            in_working = df.iloc[i]['in_working_area']
            
            if lat == 0 or lon == 0:
                trips.append(current_trip if in_trip else 0)
                continue
            
            if in_maintenance:
                if in_trip:
                    if trip_distance >= self.min_trip_distance_km and left_working_area:
                        trips.append(current_trip)
                    else:
                        trips.append(0)
                        current_trip -= 1
                    in_trip = False
                    trip_distance = 0
                    left_working_area = False
                else:
                    trips.append(0)
                continue
            
            if in_working:
                if not in_trip:
                    trips.append(0)
                elif in_trip and left_working_area:
                    if trip_distance >= self.min_trip_distance_km:
                        trips.append(current_trip)
                    else:
                        trips.append(0)
                        current_trip -= 1
                    in_trip = False
                    trip_distance = 0
                    left_working_area = False
                else:
                    trips.append(current_trip if in_trip else 0)
                continue
            
            if not in_trip:
                current_trip += 1
                in_trip = True
                left_working_area = True
                trips.append(current_trip)
            elif in_trip:
                trip_distance += df.iloc[i]['distance_km']
                trips.append(current_trip)
            else:
                trips.append(0)
        
        trip_counter = 0
        last_trip = 0
        for i in range(len(trips)):
            if trips[i] > last_trip:
                trip_counter += 1
                last_trip = trips[i]
                trips[i] = trip_counter
            elif trips[i] == last_trip:
                trips[i] = trip_counter
            else:
                trips[i] = 0
        
        df['trip'] = trips
        return df
    
    def detect_trips(self, df):
        return self.detect_real_trips(df)
    
    def detect_routes(self, df):
        valid_coords = df[(df['lat'] != 0) & (df['lon'] != 0) & (~df['in_maintenance'])]
        
        if len(valid_coords) < 10:
            return None, None
        
        coords = valid_coords[['lat', 'lon']].values
        
        try:
            eps = 0.0001
            min_samples = 30
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
            return None, None
    
    def calculate_total_hours(self, df):
        if 'time' not in df.columns or len(df) < 2:
            return 0
        
        try:
            start_time = df.iloc[0]['time']
            end_time = df.iloc[-1]['time']
            total_seconds = (end_time - start_time).total_seconds()
            total_hours = total_seconds / 3600
            
            moving_mask = (df['speed'] > 2) & (df['lat'] != 0) & (~df['in_maintenance'])
            moving_points = len(df[moving_mask])
            total_points = len(df[df['lat'] != 0])
            
            if total_points > 0:
                operating_hours = (moving_points / total_points) * total_hours
            else:
                operating_hours = total_hours
            
            return round(operating_hours, 2)
        except:
            return 0
    
    def calculate_lead_and_lift(self, df, trip_numbers):
        lead_distances = []
        lift_elevations = []
        
        if 'rl' in df.columns:
            elevation_col = 'rl'
        else:
            elevation_col = 'alt'
        
        for trip_num in trip_numbers:
            trip_df = df[df['trip'] == trip_num]
            
            if len(trip_df) < 5:
                continue
            
            working_area_points = trip_df[trip_df['in_working_area'] == True]
            if len(working_area_points) > 0:
                loading_point = working_area_points.loc[working_area_points[elevation_col].idxmin()]
            else:
                loading_point = trip_df.iloc[0]
            
            maintenance_points = trip_df[trip_df['in_maintenance'] == True]
            if len(maintenance_points) > 0:
                dumping_point = maintenance_points.loc[maintenance_points[elevation_col].idxmax()]
            else:
                dumping_point = trip_df.loc[trip_df[elevation_col].idxmax()]
            
            loading_elevation = loading_point.get(elevation_col, 0)
            dumping_elevation = dumping_point.get(elevation_col, 0)
            loading_lat = loading_point.get('lat', 0)
            loading_lon = loading_point.get('lon', 0)
            dumping_lat = dumping_point.get('lat', 0)
            dumping_lon = dumping_point.get('lon', 0)
            
            try:
                if loading_lat != 0 and loading_lon != 0 and dumping_lat != 0 and dumping_lon != 0:
                    lead_distance_km = geodesic(
                        (loading_lat, loading_lon), 
                        (dumping_lat, dumping_lon)
                    ).kilometers
                else:
                    lead_distance_km = trip_df['distance_km'].sum() / 2
            except:
                lead_distance_km = trip_df['distance_km'].sum() / 2
            
            if loading_elevation != 0 and dumping_elevation != 0:
                lift_meters = dumping_elevation - loading_elevation
            else:
                lift_meters = 0
            
            if abs(lift_meters) > 200:
                lift_meters = 100 if lift_meters > 0 else -100
            
            if lead_distance_km > 20:
                lead_distance_km = 10
            
            lead_distances.append(lead_distance_km)
            lift_elevations.append(lift_meters)
        
        avg_lead = sum(lead_distances) / len(lead_distances) if lead_distances else 0
        avg_lift = sum(lift_elevations) / len(lift_elevations) if lift_elevations else 0
        total_lead = sum(lead_distances)
        total_lift = sum(lift_elevations)
        
        return {
            'avg_lead_km': round(avg_lead, 2),
            'avg_lift_m': round(avg_lift, 1),
            'total_lead_km': round(total_lead, 2),
            'total_lift_m': round(total_lift, 1)
        }
    
    def analyze_shift_data(self, df, shift_name=""):
        if len(df) < 3:
            return None
        
        try:
            required_cols = ['lat', 'lon', 'pitch', 'speed', 'alt']
            for col in required_cols:
                if col not in df.columns:
                    df[col] = 0
            
            if 'rl' in df.columns:
                df['alt'] = df['rl']
            
            df['pitch'] = pd.to_numeric(df['pitch'], errors='coerce').fillna(0)
            df.loc[df['pitch'] > 30, 'pitch'] = df.loc[df['pitch'] > 30, 'pitch'] - 90
            df.loc[df['pitch'] < -30, 'pitch'] = df.loc[df['pitch'] < -30, 'pitch'] + 90
            
            df = self.calculate_distance_and_fuel(df)
            df = self.detect_trips(df)
            route_a, route_b = self.detect_routes(df)
            
            maintenance_records = df[df['in_maintenance'] == True]
            maintenance_minutes = len(maintenance_records) if 'time' in df.columns else 0
            total_time = len(df)
            maintenance_percentage = (maintenance_minutes / total_time * 100) if total_time > 0 else 0
            
            valid_gps = df[(df['lat'] != 0) & (df['lon'] != 0) & (df['distance_km'] > 0)]
            
            if len(valid_gps) == 0:
                return None
            
            total_distance = valid_gps['distance_km'].sum()
            total_fuel = valid_gps['fuel_l'].sum()
            total_cost = total_fuel * self.diesel_price
            
            valid_trips = []
            trip_details = []
            total_trip_distance = 0
            total_trip_fuel = 0
            slow_trips = []
            
            for trip_num in sorted(df[df['trip'] > 0]['trip'].unique()):
                if len(valid_trips) >= self.max_trips_per_shift:
                    break
                    
                trip_data = df[df['trip'] == trip_num]
                
                if len(trip_data) < 5:
                    continue
                
                trip_distance = trip_data['distance_km'].sum()
                
                if trip_distance < self.min_trip_distance_km:
                    continue
                
                trip_fuel = trip_data['fuel_l'].sum()
                trip_cost = trip_fuel * self.diesel_price
                
                total_trip_distance += trip_distance
                total_trip_fuel += trip_fuel
                valid_trips.append(trip_num)
                
                trip_duration_hours = 0
                if 'time' in trip_data.columns and len(trip_data) > 1:
                    time_diff = trip_data.iloc[-1]['time'] - trip_data.iloc[0]['time']
                    trip_duration_hours = time_diff.total_seconds() / 3600
                    
                    if trip_duration_hours > 1:
                        slow_trips.append({
                            'trip_num': int(trip_num),
                            'duration_hours': round(trip_duration_hours, 2),
                            'distance_km': round(trip_distance, 2),
                            'fuel_l': round(trip_fuel, 2)
                        })
                
                avg_speed = trip_data['speed'].mean() if 'speed' in trip_data.columns else 0
                if avg_speed > self.max_realistic_speed:
                    avg_speed = self.max_realistic_speed
                
                trip_details.append({
                    'trip_number': int(trip_num),
                    'start_time': str(trip_data.iloc[0]['time']) if 'time' in trip_data.columns else 'N/A',
                    'end_time': str(trip_data.iloc[-1]['time']) if 'time' in trip_data.columns else 'N/A',
                    'duration_hours': round(trip_duration_hours, 2),
                    'duration_minutes': round(trip_duration_hours * 60, 2),
                    'distance_km': round(float(trip_distance), 2),
                    'fuel_l': round(float(trip_fuel), 2),
                    'cost_rs': round(float(trip_cost), 2),
                    'avg_speed': round(float(avg_speed), 1)
                })
            
            unique_trips = len(valid_trips)
            total_tons = unique_trips * self.ton_per_trip
            
            lead_lift_data = self.calculate_lead_and_lift(df, valid_trips)
            
            operating_time_hours = self.calculate_total_hours(df)
            total_time_hours = 0
            
            if 'time' in df.columns and len(df) > 1:
                time_span = df.iloc[-1]['time'] - df.iloc[0]['time']
                total_time_hours = time_span.total_seconds() / 3600
            
            # Use 8 hours for shift
            total_shift_hours = 8.0
            trips_per_hour = unique_trips / total_shift_hours if total_shift_hours > 0 else 0
            
            if trips_per_hour > 1.8:
                trips_per_hour = 1.8
            
            trips_per_km = unique_trips / total_distance if total_distance > 0 else 0
            fuel_per_hour = total_fuel / total_shift_hours if total_shift_hours > 0 else 0
            fuel_per_km = total_fuel / total_distance if total_distance > 0 else 0
            fuel_per_trip = total_fuel / unique_trips if unique_trips > 0 else 0
            
            target_met = trips_per_hour >= 1.5
            target_status = "✓ Target Met (≥1.5 trips/hr)" if target_met else "✗ Target Not Met (<1.5 trips/hr)"
            
            moving = df[(df['speed'] > 2) & (df['lat'] != 0) & (~df['in_maintenance'])]
            avg_speed = float(moving['speed'].mean()) if len(moving) > 0 else 0
            if avg_speed > self.max_realistic_speed:
                avg_speed = self.max_realistic_speed
            
            gradient_stats = {}
            
            steep_up_mask = df['gradient_class'] == 'Steep Up'
            steep_up_data = df[steep_up_mask]
            steep_up_distance = steep_up_data['distance_km'].sum()
            steep_up_fuel = steep_up_data['fuel_l'].sum()
            steep_up_speed = steep_up_data['speed'].mean() if len(steep_up_data) > 0 else 0
            steep_up_time = steep_up_distance / steep_up_speed if steep_up_speed > 0 else 0
            
            steep_down_mask = df['gradient_class'] == 'Steep Down'
            steep_down_data = df[steep_down_mask]
            steep_down_distance = steep_down_data['distance_km'].sum()
            steep_down_fuel = steep_down_data['fuel_l'].sum()
            steep_down_speed = steep_down_data['speed'].mean() if len(steep_down_data) > 0 else 0
            steep_down_time = steep_down_distance / steep_down_speed if steep_down_speed > 0 else 0
            
            flat_mask = df['gradient_class'] == 'Flat'
            flat_data = df[flat_mask]
            flat_distance = flat_data['distance_km'].sum()
            flat_fuel = flat_data['fuel_l'].sum()
            flat_speed = flat_data['speed'].mean() if len(flat_data) > 0 else 0
            flat_time = flat_distance / flat_speed if flat_speed > 0 else 0
            
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
                'speed_kmhr': round(steep_up_speed, 1),
                'time_hr': round(steep_up_time, 2),
                'cost_rs': round(steep_up_fuel * self.diesel_price, 2),
                'percentage': round(steep_up_distance / total_distance * 100, 1) if total_distance > 0 else 0
            }
            
            gradient_stats['Steep Down'] = {
                'distance_km': round(steep_down_distance, 2),
                'fuel_l': round(steep_down_fuel, 2),
                'speed_kmhr': round(steep_down_speed, 1),
                'time_hr': round(steep_down_time, 2),
                'cost_rs': round(steep_down_fuel * self.diesel_price, 2),
                'percentage': round(steep_down_distance / total_distance * 100, 1) if total_distance > 0 else 0
            }
            
            gradient_stats['Flat'] = {
                'distance_km': round(flat_distance, 2),
                'fuel_l': round(flat_fuel, 2),
                'speed_kmhr': round(flat_speed, 1),
                'time_hr': round(flat_time, 2),
                'cost_rs': round(flat_fuel * self.diesel_price, 2),
                'percentage': round(flat_distance / total_distance * 100, 1) if total_distance > 0 else 0
            }
            
            idle_points = len(df[(df['speed'] <= 2) & (df['lat'] != 0) & (~df['in_maintenance'])]) if 'speed' in df.columns else 0
            total_points = len(df[(df['lat'] != 0)])
            idle_ratio = idle_points / total_points if total_points > 0 else 0
            
            if idle_ratio > 0.4:
                idle_pattern = "High Idle Time"
            elif idle_ratio > 0.2:
                idle_pattern = "Medium Idle Time"
            else:
                idle_pattern = "Low Idle Time"
            
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
            
            if fuel_per_km < 0.4:
                fuel_efficiency_rating = "Excellent"
            elif fuel_per_km < 0.65:
                fuel_efficiency_rating = "Good"
            elif fuel_per_km < 0.85:
                fuel_efficiency_rating = "Average"
            else:
                fuel_efficiency_rating = "Needs Improvement"
            
            return {
                'total_distance': round(total_distance, 1),
                'total_fuel': round(total_fuel, 1),
                'total_cost': round(total_cost, 2),
                'total_tons': total_tons,
                'cost_rs': round(total_cost, 2),
                'fuel_per_km': round(fuel_per_km, 2),
                'fuel_per_hour': round(fuel_per_hour, 1),
                'fuel_per_trip': round(fuel_per_trip, 1),
                'trips': unique_trips,
                'trips_per_hour': round(trips_per_hour, 2),
                'trips_per_km': round(trips_per_km, 3),
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
                'maintenance_time_minutes': round(maintenance_minutes, 1),
                'maintenance_time_pct': round(maintenance_percentage, 1),
                'in_maintenance': maintenance_percentage > 0,
                'target_status': target_status,
                'target_met': target_met,
                'slow_trips': slow_trips,
                'total_time_hours': round(total_time_hours, 2),
                'operating_time_hours': operating_time_hours,
                'total_hours': total_shift_hours,
                'total_shift_hours': total_shift_hours,
                'avg_lead_km': lead_lift_data.get('avg_lead_km', 0),
                'avg_lift_m': lead_lift_data.get('avg_lift_m', 0),
                'total_lead_km': lead_lift_data.get('total_lead_km', 0),
                'total_lift_m': lead_lift_data.get('total_lift_m', 0)
            }
            
        except Exception as e:
            print(f"Error: {e}", file=sys.stderr)
            return None
    
    def analyze_daily_data(self, df):
        if len(df) < 3:
            return None
        
        try:
            required_cols = ['lat', 'lon', 'pitch', 'speed', 'alt']
            for col in required_cols:
                if col not in df.columns:
                    df[col] = 0
            
            if 'rl' in df.columns:
                df['alt'] = df['rl']
            
            df['pitch'] = pd.to_numeric(df['pitch'], errors='coerce').fillna(0)
            df.loc[df['pitch'] > 30, 'pitch'] = df.loc[df['pitch'] > 30, 'pitch'] - 90
            df.loc[df['pitch'] < -30, 'pitch'] = df.loc[df['pitch'] < -30, 'pitch'] + 90
            
            df = self.calculate_distance_and_fuel(df)
            df = self.detect_trips(df)
            route_a, route_b = self.detect_routes(df)
            
            maintenance_records = df[df['in_maintenance'] == True]
            maintenance_minutes = len(maintenance_records) if 'time' in df.columns else 0
            total_time = len(df)
            maintenance_percentage = (maintenance_minutes / total_time * 100) if total_time > 0 else 0
            
            valid_gps = df[(df['lat'] != 0) & (df['lon'] != 0) & (df['distance_km'] > 0)]
            
            if len(valid_gps) == 0:
                return None
            
            total_distance = valid_gps['distance_km'].sum()
            total_fuel = valid_gps['fuel_l'].sum()
            total_cost = total_fuel * self.diesel_price
            
            valid_trips = []
            trip_details = []
            total_trip_distance = 0
            total_trip_fuel = 0
            slow_trips = []
            
            for trip_num in sorted(df[df['trip'] > 0]['trip'].unique()):
                trip_data = df[df['trip'] == trip_num]
                
                if len(trip_data) < 5:
                    continue
                
                trip_distance = trip_data['distance_km'].sum()
                
                if trip_distance < self.min_trip_distance_km:
                    continue
                
                trip_fuel = trip_data['fuel_l'].sum()
                trip_cost = trip_fuel * self.diesel_price
                
                total_trip_distance += trip_distance
                total_trip_fuel += trip_fuel
                valid_trips.append(trip_num)
                
                trip_duration_hours = 0
                if 'time' in trip_data.columns and len(trip_data) > 1:
                    time_diff = trip_data.iloc[-1]['time'] - trip_data.iloc[0]['time']
                    trip_duration_hours = time_diff.total_seconds() / 3600
                    
                    if trip_duration_hours > 1.5:
                        slow_trips.append({
                            'trip_num': int(trip_num),
                            'duration_hours': round(trip_duration_hours, 2),
                            'distance_km': round(trip_distance, 2),
                            'fuel_l': round(trip_fuel, 2)
                        })
                
                avg_speed = trip_data['speed'].mean() if 'speed' in trip_data.columns else 0
                if avg_speed > self.max_realistic_speed:
                    avg_speed = self.max_realistic_speed
                
                trip_details.append({
                    'trip_number': int(trip_num),
                    'start_time': str(trip_data.iloc[0]['time']) if 'time' in trip_data.columns else 'N/A',
                    'end_time': str(trip_data.iloc[-1]['time']) if 'time' in trip_data.columns else 'N/A',
                    'duration_hours': round(trip_duration_hours, 2),
                    'duration_minutes': round(trip_duration_hours * 60, 2),
                    'distance_km': round(float(trip_distance), 2),
                    'fuel_l': round(float(trip_fuel), 2),
                    'cost_rs': round(float(trip_cost), 2),
                    'avg_speed': round(float(avg_speed), 1)
                })
            
            unique_trips = len(valid_trips)
            total_tons = unique_trips * self.ton_per_trip
            
            lead_lift_data = self.calculate_lead_and_lift(df, valid_trips)
            
            operating_time_hours = self.calculate_total_hours(df)
            total_time_hours = 0
            
            if 'time' in df.columns and len(df) > 1:
                time_span = df.iloc[-1]['time'] - df.iloc[0]['time']
                total_time_hours = time_span.total_seconds() / 3600
            
            # Use 24 hours for daily
            total_day_hours = 24.0
            trips_per_hour = unique_trips / total_day_hours if total_day_hours > 0 else 0
            
            if trips_per_hour > 1.5:
                trips_per_hour = 1.5
            
            trips_per_km = unique_trips / total_distance if total_distance > 0 else 0
            fuel_per_hour = total_fuel / total_day_hours if total_day_hours > 0 else 0
            fuel_per_km = total_fuel / total_distance if total_distance > 0 else 0
            fuel_per_trip = total_fuel / unique_trips if unique_trips > 0 else 0
            
            target_met = trips_per_hour >= 1.5
            
            moving = df[(df['speed'] > 2) & (df['lat'] != 0) & (~df['in_maintenance'])]
            avg_speed = float(moving['speed'].mean()) if len(moving) > 0 else 0
            if avg_speed > self.max_realistic_speed:
                avg_speed = self.max_realistic_speed
            
            gradient_stats = {}
            
            steep_up_mask = df['gradient_class'] == 'Steep Up'
            steep_up_data = df[steep_up_mask]
            steep_up_distance = steep_up_data['distance_km'].sum()
            steep_up_fuel = steep_up_data['fuel_l'].sum()
            steep_up_speed = steep_up_data['speed'].mean() if len(steep_up_data) > 0 else 0
            steep_up_time = steep_up_distance / steep_up_speed if steep_up_speed > 0 else 0
            
            steep_down_mask = df['gradient_class'] == 'Steep Down'
            steep_down_data = df[steep_down_mask]
            steep_down_distance = steep_down_data['distance_km'].sum()
            steep_down_fuel = steep_down_data['fuel_l'].sum()
            steep_down_speed = steep_down_data['speed'].mean() if len(steep_down_data) > 0 else 0
            steep_down_time = steep_down_distance / steep_down_speed if steep_down_speed > 0 else 0
            
            flat_mask = df['gradient_class'] == 'Flat'
            flat_data = df[flat_mask]
            flat_distance = flat_data['distance_km'].sum()
            flat_fuel = flat_data['fuel_l'].sum()
            flat_speed = flat_data['speed'].mean() if len(flat_data) > 0 else 0
            flat_time = flat_distance / flat_speed if flat_speed > 0 else 0
            
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
                'speed_kmhr': round(steep_up_speed, 1),
                'time_hr': round(steep_up_time, 2)
            }
            
            gradient_stats['Steep Down'] = {
                'distance_km': round(steep_down_distance, 2),
                'fuel_l': round(steep_down_fuel, 2),
                'speed_kmhr': round(steep_down_speed, 1),
                'time_hr': round(steep_down_time, 2)
            }
            
            gradient_stats['Flat'] = {
                'distance_km': round(flat_distance, 2),
                'fuel_l': round(flat_fuel, 2),
                'speed_kmhr': round(flat_speed, 1),
                'time_hr': round(flat_time, 2)
            }
            
            idle_points = len(df[(df['speed'] <= 2) & (df['lat'] != 0) & (~df['in_maintenance'])]) if 'speed' in df.columns else 0
            total_points = len(df[(df['lat'] != 0)])
            idle_ratio = idle_points / total_points if total_points > 0 else 0
            
            if idle_ratio > 0.4:
                idle_pattern = "High Idle Time"
            elif idle_ratio > 0.2:
                idle_pattern = "Medium Idle Time"
            else:
                idle_pattern = "Low Idle Time"
            
            steep_up_pct = steep_up_distance / total_distance * 100 if total_distance > 0 else 0
            
            if steep_up_pct > 30:
                route_type = "Steep Route (High Grade)"
            elif steep_up_pct > 15:
                route_type = "Mixed Terrain"
            else:
                route_type = "Gentle Route"
            
            if fuel_per_km < 0.4:
                fuel_efficiency_rating = "Excellent"
            elif fuel_per_km < 0.65:
                fuel_efficiency_rating = "Good"
            elif fuel_per_km < 0.85:
                fuel_efficiency_rating = "Average"
            else:
                fuel_efficiency_rating = "Needs Improvement"
            
            return {
                'total_distance': round(total_distance, 1),
                'total_fuel': round(total_fuel, 1),
                'total_cost': round(total_cost, 2),
                'total_tons': total_tons,
                'fuel_per_km': round(fuel_per_km, 2),
                'fuel_per_hour': round(fuel_per_hour, 1),
                'fuel_per_trip': round(fuel_per_trip, 1),
                'trips': unique_trips,
                'trips_per_hour': round(trips_per_hour, 2),
                'trips_per_km': round(trips_per_km, 3),
                'avg_speed': round(avg_speed, 1),
                'fuel_efficiency_rating': fuel_efficiency_rating,
                'gradient_stats': gradient_stats,
                'idle_pattern': idle_pattern,
                'idle_ratio': round(idle_ratio * 100, 1),
                'route_type': route_type,
                'trip_details': trip_details[:50],
                'route_a': route_a,
                'route_b': route_b,
                'maintenance_time_minutes': round(maintenance_minutes, 1),
                'maintenance_time_pct': round(maintenance_percentage, 1),
                'target_met': target_met,
                'total_time_hours': round(total_time_hours, 2),
                'operating_time_hours': operating_time_hours,
                'total_hours': total_day_hours,
                'total_day_hours': total_day_hours,
                'avg_lead_km': lead_lift_data.get('avg_lead_km', 0),
                'avg_lift_m': lead_lift_data.get('avg_lift_m', 0),
                'total_lead_km': lead_lift_data.get('total_lead_km', 0),
                'total_lift_m': lead_lift_data.get('total_lift_m', 0)
            }
            
        except Exception as e:
            print(f"Error: {e}", file=sys.stderr)
            return None
    
    def analyze_monthly_data(self, df):
        if len(df) < 3:
            return None
        
        try:
            if 'time' not in df.columns:
                return self.analyze_daily_data(df)
            
            df['date'] = df['time'].dt.date
            unique_dates = df['date'].unique()
            
            if len(unique_dates) == 0:
                return self.analyze_daily_data(df)
            
            print(f"Analyzing {len(unique_dates)} days for monthly report", file=sys.stderr)
            
            daily_results = []
            
            for date in unique_dates:
                daily_df = df[df['date'] == date].copy()
                
                if len(daily_df) < 10:
                    continue
                
                daily_result = self.analyze_daily_data(daily_df)
                
                if daily_result:
                    daily_results.append(daily_result)
            
            if not daily_results:
                return self.analyze_daily_data(df)
            
            total_trips = sum(r.get('trips', 0) for r in daily_results)
            total_distance = sum(r.get('total_distance', 0) for r in daily_results)
            total_fuel = sum(r.get('total_fuel', 0) for r in daily_results)
            total_cost = sum(r.get('total_cost', 0) for r in daily_results)
            total_days = len(daily_results)
            
            total_shift_hours = total_days * 8.0
            trips_per_hour = total_trips / total_shift_hours if total_shift_hours > 0 else 0
            
            if trips_per_hour > 1.8:
                trips_per_hour = 1.8
            
            avg_fuel_per_km = total_fuel / total_distance if total_distance > 0 else 0
            avg_fuel_per_hour = total_fuel / total_shift_hours if total_shift_hours > 0 else 0
            avg_fuel_per_trip = total_fuel / total_trips if total_trips > 0 else 0
            avg_trips_per_km = total_trips / total_distance if total_distance > 0 else 0
            
            total_distance_weighted_speed = sum(r.get('total_distance', 0) * r.get('avg_speed', 0) for r in daily_results)
            avg_speed = total_distance_weighted_speed / total_distance if total_distance > 0 else 0
            
            gradient_stats = {
                'Steep Up': {'distance_km': 0, 'fuel_l': 0, 'time_hr': 0},
                'Steep Down': {'distance_km': 0, 'fuel_l': 0, 'time_hr': 0},
                'Flat': {'distance_km': 0, 'fuel_l': 0, 'time_hr': 0}
            }
            
            for r in daily_results:
                grad_stats = r.get('gradient_stats', {})
                for grad_type in ['Steep Up', 'Steep Down', 'Flat']:
                    if grad_type in grad_stats:
                        gradient_stats[grad_type]['distance_km'] += grad_stats[grad_type].get('distance_km', 0)
                        gradient_stats[grad_type]['fuel_l'] += grad_stats[grad_type].get('fuel_l', 0)
                        gradient_stats[grad_type]['time_hr'] += grad_stats[grad_type].get('time_hr', 0)
            
            for grad_type in gradient_stats:
                dist = gradient_stats[grad_type]['distance_km']
                fuel = gradient_stats[grad_type]['fuel_l']
                time_hr = gradient_stats[grad_type]['time_hr']
                
                gradient_stats[grad_type]['speed_kmhr'] = round(dist / time_hr, 1) if time_hr > 0 else 0
                gradient_stats[grad_type]['cost_rs'] = round(fuel * self.diesel_price, 2)
                gradient_stats[grad_type]['percentage'] = round(dist / total_distance * 100, 1) if total_distance > 0 else 0
            
            avg_idle_ratio = sum(r.get('idle_ratio', 0) for r in daily_results) / len(daily_results) if daily_results else 0
            
            if avg_idle_ratio > 40:
                idle_pattern = "High Idle Time"
            elif avg_idle_ratio > 20:
                idle_pattern = "Medium Idle Time"
            else:
                idle_pattern = "Low Idle Time"
            
            steep_up_pct = gradient_stats['Steep Up']['percentage']
            if steep_up_pct > 30:
                route_type = "Steep Route (High Grade)"
            elif steep_up_pct > 15:
                route_type = "Mixed Terrain"
            else:
                route_type = "Gentle Route"
            
            if avg_fuel_per_km < 0.4:
                fuel_efficiency_rating = "Excellent"
            elif avg_fuel_per_km < 0.65:
                fuel_efficiency_rating = "Good"
            elif avg_fuel_per_km < 0.85:
                fuel_efficiency_rating = "Average"
            else:
                fuel_efficiency_rating = "Needs Improvement"
            
            target_met = trips_per_hour >= 1.5
            
            all_trip_details = []
            for r in daily_results:
                all_trip_details.extend(r.get('trip_details', []))
                if len(all_trip_details) >= 50:
                    break
            
            route_a = None
            route_b = None
            for r in daily_results:
                if r.get('route_a'):
                    route_a = r.get('route_a')
                    route_b = r.get('route_b')
                    break
            
            avg_lead_km = sum(r.get('avg_lead_km', 0) for r in daily_results) / len(daily_results) if daily_results else 0
            avg_lift_m = sum(r.get('avg_lift_m', 0) for r in daily_results) / len(daily_results) if daily_results else 0
            total_lead_km = sum(r.get('total_lead_km', 0) for r in daily_results)
            total_lift_m = sum(r.get('total_lift_m', 0) for r in daily_results)
            
            return {
                'total_distance': round(total_distance, 1),
                'total_fuel': round(total_fuel, 1),
                'total_cost': round(total_cost, 2),
                'total_tons': total_trips * self.ton_per_trip,
                'cost_rs': round(total_cost, 2),
                'fuel_per_km': round(avg_fuel_per_km, 2),
                'fuel_per_hour': round(avg_fuel_per_hour, 1),
                'fuel_per_trip': round(avg_fuel_per_trip, 1),
                'trips': total_trips,
                'trips_per_hour': round(trips_per_hour, 2),
                'trips_per_km': round(avg_trips_per_km, 4),
                'avg_speed': round(avg_speed, 1),
                'fuel_efficiency_rating': fuel_efficiency_rating,
                'gradient_stats': gradient_stats,
                'idle_pattern': idle_pattern,
                'idle_ratio': round(avg_idle_ratio, 1),
                'route_type': route_type,
                'trip_details': all_trip_details[:50],
                'route_a': route_a,
                'route_b': route_b,
                'target_met': target_met,
                'total_shift_hours': round(total_shift_hours, 2),
                'total_hours': round(total_shift_hours, 2),
                'number_of_days': total_days,
                'avg_lead_km': round(avg_lead_km, 2),
                'avg_lift_m': round(avg_lift_m, 1),
                'total_lead_km': round(total_lead_km, 2),
                'total_lift_m': round(total_lift_m, 1)
            }
            
        except Exception as e:
            print(f"Error in analyze_monthly_data: {e}", file=sys.stderr)
            import traceback
            traceback.print_exc(file=sys.stderr)
            return self.analyze_daily_data(df)
    
    def analyze_excavator_data(self, df):
        if len(df) < 3:
            return None
        
        try:
            df = self.calculate_distance_and_fuel(df)
            
            maintenance_records = df[df['in_maintenance'] == True]
            maintenance_minutes = len(maintenance_records) if 'time' in df.columns else 0
            total_time = len(df)
            maintenance_percentage = (maintenance_minutes / total_time * 100) if total_time > 0 else 0
            
            valid_gps = df[(df['lat'] != 0) & (df['lon'] != 0) & (df['distance_km'] > 0)]
            
            if len(valid_gps) == 0:
                return None
            
            total_distance = valid_gps['distance_km'].sum()
            total_fuel = valid_gps['fuel_l'].sum()
            total_cost = total_fuel * self.diesel_price
            
            idle_points = len(df[(df['speed'] <= 2) & (df['lat'] != 0) & (~df['in_maintenance'])]) if 'speed' in df.columns else 0
            total_points = len(df[(df['lat'] != 0)])
            idle_ratio = idle_points / total_points if total_points > 0 else 0
            idle_hours = idle_ratio * (total_time / 60) if 'time' in df.columns else 0
            running_hours = (1 - idle_ratio) * (total_time / 60) if 'time' in df.columns else 0
            
            operating_status = "Stationary Operation" if total_distance < 5 else "Moving Operation"
            fuel_per_hour = total_fuel / (running_hours + idle_hours) if (running_hours + idle_hours) > 0 else 0
            
            return {
                'total_distance': round(total_distance, 1),
                #'total_fuel': round(total_fuel, 1),
                #'total_cost': round(total_cost, 2),
                #'fuel_per_km': round(total_fuel / total_distance, 2) if total_distance > 0 else 0,
                #'fuel_per_hour': round(fuel_per_hour, 1),
                'idle_hours': round(idle_hours, 2),
                'running_hours': round(running_hours, 2),
                'idle_ratio': round(idle_ratio * 100, 1),
                'maintenance_time_pct': round(maintenance_percentage, 1),
                'operating_status': operating_status
            }
            
        except Exception as e:
            return None
    
    def analyze_bulldozer_data(self, df):
        if len(df) < 3:
            return None
        
        try:
            df = self.calculate_distance_and_fuel(df)
            
            total_hours = 0
            operating_hours = 0
            
            if 'time' in df.columns and len(df) > 1:
                time_span = df.iloc[-1]['time'] - df.iloc[0]['time']
                total_hours = time_span.total_seconds() / 3600
                total_points = len(df)
                if total_points > 0:
                    moving_points = len(df[(df['speed'] > 2) & (df['lat'] != 0) & (~df['in_maintenance'])])
                    operating_hours = (moving_points / total_points) * total_hours
            
            valid_gps = df[(df['lat'] != 0) & (df['lon'] != 0) & (df['distance_km'] > 0)]
            
            total_distance = valid_gps['distance_km'].sum() if len(valid_gps) > 0 else 0
            total_fuel = valid_gps['fuel_l'].sum() if len(valid_gps) > 0 else 0
            total_cost = total_fuel * self.diesel_price
            
            fuel_per_hour = total_fuel / operating_hours if operating_hours > 0 else 0
            
            if operating_hours > total_hours * 0.3:
                status = "Active Operation"
            else:
                status = "Standby/Idle"
            
            return {
                'total_distance': round(total_distance, 1),
                'total_fuel': round(total_fuel, 1),
                'total_cost': round(total_cost, 2),
                'operating_hours': round(operating_hours, 2),
                'fuel_per_hour': round(fuel_per_hour, 2),
                'status': status
            }
            
        except Exception as e:
            return None


class ExcelReportGenerator:
    def __init__(self):
        self.wb = Workbook()
        self.setup_styles()
        self.analyzer = SimpleMiningAnalytics()
    
    def setup_styles(self):
        self.header_fill = PatternFill(start_color='2C3E50', end_color='2C3E50', fill_type='solid')
        self.header_font = Font(color='FFFFFF', bold=True, size=11)
        self.centered = Alignment(horizontal='center', vertical='center')
        self.right_aligned = Alignment(horizontal='right', vertical='center')
        self.left_aligned = Alignment(horizontal='left', vertical='center')
        self.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        self.good_fill = PatternFill(start_color='2ECC71', end_color='2ECC71', fill_type='solid')
        self.bad_fill = PatternFill(start_color='E74C3C', end_color='E74C3C', fill_type='solid')
        self.warning_fill = PatternFill(start_color='F39C12', end_color='F39C12', fill_type='solid')
    
    def create_hauler_summary_sheet(self, hauler_results):
        ws = self.wb.create_sheet("Hauler Summary")
        
        headers = ['HAULER ID', 'Total Trips', 'Tons/Trips', 'Trips/Hr', 'Trips/Km', 'Km/Hr', 'Fuel/Hr', 'Fuel/Km', 'Fuel/Trip', 
                   'Avg Lead (km)', 'Avg Lift (m)', 'Total Fuel', 'Total Km']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.centered
            cell.border = self.border
        
        row = 2
        for device_id, data in sorted(hauler_results.items(), key=lambda x: x[1].get('total_fuel', 0), reverse=True):
            hauler_id = self.analyzer.map_device_id(device_id)
            trips = data.get('trips', 0)
            tons_per_trip = data.get('total_tons', trips * 35) / trips if trips > 0 else 0
            
            ws.cell(row=row, column=1, value=hauler_id)
            ws.cell(row=row, column=2, value=trips)
            ws.cell(row=row, column=3, value=round(tons_per_trip, 1))
            ws.cell(row=row, column=4, value=data.get('trips_per_hour', 0))
            ws.cell(row=row, column=5, value=data.get('trips_per_km', 0))
            ws.cell(row=row, column=6, value=data.get('avg_speed', 0))
            ws.cell(row=row, column=7, value=data.get('fuel_per_hour', 0))
            ws.cell(row=row, column=8, value=data.get('fuel_per_km', 0))
            ws.cell(row=row, column=9, value=data.get('fuel_per_trip', 0))
            ws.cell(row=row, column=10, value=data.get('avg_lead_km', 0))
            ws.cell(row=row, column=11, value=data.get('avg_lift_m', 0))
            ws.cell(row=row, column=12, value=data.get('total_fuel', 0))
            ws.cell(row=row, column=13, value=data.get('total_distance', 0))
            
            trips_per_hour_cell = ws.cell(row=row, column=4)
            if data.get('trips_per_hour', 0) >= 1.5:
                trips_per_hour_cell.fill = self.good_fill
            elif data.get('trips_per_hour', 0) >= 1.0:
                trips_per_hour_cell.fill = self.warning_fill
            else:
                trips_per_hour_cell.fill = self.bad_fill
            
            for col in range(1, 14):
                ws.cell(row=row, column=col).border = self.border
                if col > 1:
                    ws.cell(row=row, column=col).alignment = self.right_aligned
                else:
                    ws.cell(row=row, column=col).alignment = self.centered
            
            row += 1
        
        if hauler_results:
            total_trips = sum(v.get('trips', 0) for v in hauler_results.values())
            total_fuel = sum(v.get('total_fuel', 0) for v in hauler_results.values())
            total_distance = sum(v.get('total_distance', 0) for v in hauler_results.values())
            
            ws.cell(row=row, column=1, value="TOTAL").font = Font(bold=True)
            ws.cell(row=row, column=2, value=total_trips)
            ws.cell(row=row, column=3, value=round(total_trips * 35 / len(hauler_results), 1) if hauler_results else 0)
            ws.cell(row=row, column=12, value=round(total_fuel, 1))
            ws.cell(row=row, column=13, value=round(total_distance, 1))
            
            for col in range(1, 14):
                ws.cell(row=row, column=col).border = self.border
                if col > 1:
                    ws.cell(row=row, column=col).alignment = self.right_aligned
        
        column_widths = [12, 10, 10, 10, 10, 10, 10, 10, 12, 14, 12, 12, 12]
        for col, width in enumerate(column_widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = width
        
        return ws
    
    def create_gradient_analysis_sheet(self, hauler_results):
        ws = self.wb.create_sheet("Gradient Analysis")
        
        headers = ['Hauler Id', 'Steep Up (Km)', 'Steep Up Fuel(l)', 'Steep Up Speed (Km/hr)', 'Steep up Time',
                   'Flat (Km)', 'Flat Fuel(l)', 'Flat Speed (Km/hr)', 'Flat Time',
                   'Steep Down (Km)', 'Steep Down Fuel (l)', 'Steep Down Speed (Km/hr)', 'Steep Down Time']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.centered
            cell.border = self.border
        
        row = 2
        for device_id, data in sorted(hauler_results.items(), key=lambda x: x[1].get('total_fuel', 0), reverse=True):
            hauler_id = self.analyzer.map_device_id(device_id)
            grad_stats = data.get('gradient_stats', {})
            
            steep_up = grad_stats.get('Steep Up', {})
            steep_down = grad_stats.get('Steep Down', {})
            flat = grad_stats.get('Flat', {})
            
            ws.cell(row=row, column=1, value=hauler_id)
            ws.cell(row=row, column=2, value=steep_up.get('distance_km', 0))
            ws.cell(row=row, column=3, value=steep_up.get('fuel_l', 0))
            ws.cell(row=row, column=4, value=steep_up.get('speed_kmhr', 0))
            ws.cell(row=row, column=5, value=steep_up.get('time_hr', 0))
            ws.cell(row=row, column=6, value=flat.get('distance_km', 0))
            ws.cell(row=row, column=7, value=flat.get('fuel_l', 0))
            ws.cell(row=row, column=8, value=flat.get('speed_kmhr', data.get('avg_speed', 0)))
            ws.cell(row=row, column=9, value=flat.get('time_hr', 0))
            ws.cell(row=row, column=10, value=steep_down.get('distance_km', 0))
            ws.cell(row=row, column=11, value=steep_down.get('fuel_l', 0))
            ws.cell(row=row, column=12, value=steep_down.get('speed_kmhr', 0))
            ws.cell(row=row, column=13, value=steep_down.get('time_hr', 0))
            
            for col in range(1, 14):
                ws.cell(row=row, column=col).border = self.border
                if col > 1:
                    ws.cell(row=row, column=col).alignment = self.right_aligned
                else:
                    ws.cell(row=row, column=col).alignment = self.centered
            
            row += 1
        
        column_widths = [12, 14, 16, 20, 14, 12, 14, 20, 14, 14, 18, 20, 16]
        for col, width in enumerate(column_widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = width
        
        return ws
    
    def create_route_taken_analysis_sheet(self, hauler_results):
        ws = self.wb.create_sheet("Route Taken Analysis")
        
        headers = ['Hauler Id', 'Routes', 'Distance (Km)', 'Fuel(L)', 'Time(Hr)', 
                   'Gradient Impact', 'Flat Segment', '% of Travel', 'Recommendation']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.centered
            cell.border = self.border
        
        row = 2
        for device_id, data in sorted(hauler_results.items(), key=lambda x: x[1].get('total_fuel', 0), reverse=True):
            hauler_id = self.analyzer.map_device_id(device_id)
            route_a = data.get('route_a')
            route_b = data.get('route_b')
            
            if route_a:
                avg_pitch = route_a.get('avg_pitch', 0)
                if avg_pitch > 8:
                    gradient_impact = "High - Steep Up"
                elif avg_pitch > 3:
                    gradient_impact = "Medium - Mild Up"
                elif avg_pitch > -3:
                    gradient_impact = "Low - Flat"
                else:
                    gradient_impact = "Low - Downhill"
                
                if route_b:
                    if route_a.get('fuel_per_km', 1) < route_b.get('fuel_per_km', 2):
                        recommendation = "Route A is more fuel efficient"
                    else:
                        recommendation = "Consider Route B for better efficiency"
                else:
                    recommendation = "Only route detected"
                
                ws.cell(row=row, column=1, value=hauler_id)
                ws.cell(row=row, column=2, value="Route A")
                ws.cell(row=row, column=3, value=route_a.get('distance_km', 0))
                ws.cell(row=row, column=4, value=route_a.get('fuel_l', 0))
                ws.cell(row=row, column=5, value=round(route_a.get('distance_km', 0) / 25, 2))
                ws.cell(row=row, column=6, value=gradient_impact)
                ws.cell(row=row, column=7, value=route_a.get('grade', 'Unknown'))
                ws.cell(row=row, column=8, value=route_a.get('percentage_of_travel', 0))
                ws.cell(row=row, column=9, value=recommendation)
                
                for col in range(1, 10):
                    ws.cell(row=row, column=col).border = self.border
                
                row += 1
            
            if route_b:
                avg_pitch = route_b.get('avg_pitch', 0)
                if avg_pitch > 8:
                    gradient_impact = "High - Steep Up"
                elif avg_pitch > 3:
                    gradient_impact = "Medium - Mild Up"
                elif avg_pitch > -3:
                    gradient_impact = "Low - Flat"
                else:
                    gradient_impact = "Low - Downhill"
                
                ws.cell(row=row, column=1, value=hauler_id)
                ws.cell(row=row, column=2, value="Route B")
                ws.cell(row=row, column=3, value=route_b.get('distance_km', 0))
                ws.cell(row=row, column=4, value=route_b.get('fuel_l', 0))
                ws.cell(row=row, column=5, value=round(route_b.get('distance_km', 0) / 25, 2))
                ws.cell(row=row, column=6, value=gradient_impact)
                ws.cell(row=row, column=7, value=route_b.get('grade', 'Unknown'))
                ws.cell(row=row, column=8, value=route_b.get('percentage_of_travel', 0))
                ws.cell(row=row, column=9, value="")
                
                for col in range(1, 10):
                    ws.cell(row=row, column=col).border = self.border
                
                row += 1
        
        column_widths = [12, 12, 14, 12, 10, 20, 14, 12, 35]
        for col, width in enumerate(column_widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = width
        
        return ws
    
    def create_excavator_summary_sheet(self, excavator_results):
        ws = self.wb.create_sheet("Excavator Summary")
        
        headers = ['Excavator Id', 'Cycle Time/Hauler', 'Tons/Hr', 'Fuel/Hr', 'Idle Time(Hr)', 
                   'Running Time(Hr)', 'Number of Passes', 'Current Fuel Left(L)', 'Distance (Km)']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.centered
            cell.border = self.border
        
        row = 2
        for device_id, data in excavator_results.items():
            excavator_id = self.analyzer.map_device_id(device_id)
            
            ws.cell(row=row, column=1, value=excavator_id)
            ws.cell(row=row, column=2, value=0)
            ws.cell(row=row, column=3, value=0)
            ws.cell(row=row, column=4, value=data.get('fuel_per_hour', 0))
            ws.cell(row=row, column=5, value=data.get('idle_hours', 0))
            ws.cell(row=row, column=6, value=data.get('running_hours', 0))
            ws.cell(row=row, column=7, value=0)
            ws.cell(row=row, column=8, value=0)
            ws.cell(row=row, column=9, value=data.get('total_distance', 0))
            
            for col in range(1, 10):
                ws.cell(row=row, column=col).border = self.border
                if col > 1:
                    ws.cell(row=row, column=col).alignment = self.right_aligned
                else:
                    ws.cell(row=row, column=col).alignment = self.centered
            
            row += 1
        
        column_widths = [15, 16, 12, 12, 14, 16, 16, 18, 12]
        for col, width in enumerate(column_widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = width
        
        return ws
    
    def create_bulldozer_analysis_sheet(self, bulldozer_results):
        ws = self.wb.create_sheet("Bulldozer Analysis")
        
        headers = ['Bulldozer Id', 'Distance (km)', 'Fuel (L)', 'Cost (₹)', 'Operating Hrs', 'Fuel per Hour (L)', 'Status']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.centered
            cell.border = self.border
        
        row = 2
        for device_id, data in bulldozer_results.items():
            ws.cell(row=row, column=1, value=self.analyzer.map_device_id(device_id))
            ws.cell(row=row, column=2, value=data.get('total_distance', 0))
            ws.cell(row=row, column=3, value=data.get('total_fuel', 0))
            ws.cell(row=row, column=4, value=data.get('total_cost', 0))
            ws.cell(row=row, column=5, value=data.get('operating_hours', 0))
            ws.cell(row=row, column=6, value=data.get('fuel_per_hour', 0))
            ws.cell(row=row, column=7, value=data.get('status', 'Unknown'))
            
            if data.get('status') == "Active Operation":
                ws.cell(row=row, column=7).fill = self.good_fill
            else:
                ws.cell(row=row, column=7).fill = self.warning_fill
            
            for col in range(1, 8):
                ws.cell(row=row, column=col).border = self.border
                if col > 1:
                    ws.cell(row=row, column=col).alignment = self.right_aligned
            row += 1
        
        for col in range(1, 8):
            ws.column_dimensions[get_column_letter(col)].width = 16
        
        return ws
    
    def create_executive_summary(self, hauler_results, excavator_results, bulldozer_results, report_type):
        ws = self.wb.create_sheet("Executive Summary", 0)
        
        title_cell = ws.cell(row=1, column=1, value="MINING HAULER PERFORMANCE DASHBOARD")
        title_cell.font = Font(size=16, bold=True, color='2C3E50')
        title_cell.alignment = self.centered
        ws.merge_cells('A1:H1')
        
        date_cell = ws.cell(row=2, column=1, value=f"Report Generated: {datetime.now().strftime('%d %B %Y %H:%M')}")
        date_cell.font = Font(size=10, italic=True)
        ws.merge_cells('A2:H2')
        
        report_type_cell = ws.cell(row=3, column=1, value=f"Report Type: {report_type.upper()} Analysis")
        report_type_cell.font = Font(size=10, bold=True, color='2C3E50')
        ws.merge_cells('A3:H3')
        
        total_fuel = sum(r.get('total_fuel', 0) for r in hauler_results.values())
        total_cost = sum(r.get('total_cost', 0) for r in hauler_results.values())
        total_distance = sum(r.get('total_distance', 0) for r in hauler_results.values())
        total_trips = sum(r.get('trips', 0) for r in hauler_results.values())
        total_hours = sum(r.get('total_hours', r.get('total_shift_hours', 0)) for r in hauler_results.values())
        target_met_count = sum(1 for r in hauler_results.values() if r.get('target_met', False))
        
        metrics = [
            ['Total Fuel Consumed', f"{total_fuel:,.1f} Liters", f"₹{total_cost:,.2f}"],
            ['Total Distance Traveled', f"{total_distance:,.1f} km", ''],
            ['Total Trips Completed', f"{total_trips}", ''],
            ['Total Operating Hours', f"{total_hours:,.1f} hours", ''],
            ['Number of Haulers', f"{len(hauler_results)}", ''],
            ['Number of Excavators', f"{len(excavator_results)}", ''],
            ['Number of Bulldozers', f"{len(bulldozer_results)}", ''],
            ['', '', ''],
            ['TARGET SUMMARY (1.5 trips per hour)', '', ''],
            ['Haulers Meeting Target', f"{target_met_count} / {len(hauler_results)}", f"{target_met_count/len(hauler_results)*100:.0f}%" if hauler_results else '0%'],
        ]
        
        row = 5
        for metric in metrics:
            ws.cell(row=row, column=1, value=metric[0]).font = Font(bold=True)
            ws.cell(row=row, column=2, value=metric[1])
            if len(metric) > 2:
                ws.cell(row=row, column=3, value=metric[2])
            row += 1
        
        for col in range(1, 4):
            ws.column_dimensions[get_column_letter(col)].width = 30
        
        return ws
    
    def generate_report(self, hauler_results, excavator_results, bulldozer_results, report_type="shift"):
        if 'Sheet' in self.wb.sheetnames:
            std = self.wb['Sheet']
            self.wb.remove(std)
        
        self.create_executive_summary(hauler_results, excavator_results, bulldozer_results, report_type)
        self.create_hauler_summary_sheet(hauler_results)
        self.create_gradient_analysis_sheet(hauler_results)
        self.create_route_taken_analysis_sheet(hauler_results)
        
        if excavator_results:
            self.create_excavator_summary_sheet(excavator_results)
        
        if bulldozer_results:
            self.create_bulldozer_analysis_sheet(bulldozer_results)
        
        return self.wb


def main():
    try:
        input_data = sys.stdin.read()
        if not input_data:
            print(json.dumps({'status': 'error', 'error': 'No input data'}))
            return
        
        params = json.loads(input_data)
        df = pd.DataFrame(params.get('data', []))
        report_type = params.get('report_type', 'shift')
        
        if df.empty:
            print(json.dumps({
                'status': 'success',
                'report': '',
                'filename': 'No_Data.xlsx',
                'devices_discovered': [],
                'record_count': 0
            }))
            return
        
        print(f"Total records: {len(df)}", file=sys.stderr)
        print(f"Report type: {report_type}", file=sys.stderr)
        
        essential_cols = ['lat', 'lon', 'pitch', 'speed']
        for col in essential_cols:
            if col not in df.columns:
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
        
        if 'rl' in df.columns:
            df['rl'] = pd.to_numeric(df['rl'], errors='coerce').fillna(0)
            print(f"RL column found - Range: {df['rl'].min()} to {df['rl'].max()}", file=sys.stderr)
        
        analyzer = SimpleMiningAnalytics()
        
        hauler_results = {}
        excavator_results = {}
        bulldozer_results = {}
        
        for device_id in df['device_id'].unique():
            if pd.isna(device_id):
                continue
            
            device_df = df[df['device_id'] == device_id].copy()
            
            if len(device_df) < 3:
                continue
            
            if analyzer.is_hauler(device_id):
                if report_type == 'shift':
                    result = analyzer.analyze_shift_data(device_df)
                elif report_type == 'daily':
                    result = analyzer.analyze_daily_data(device_df)
                else:
                    result = analyzer.analyze_monthly_data(device_df)
                
                if result:
                    hauler_results[str(device_id)] = result
                    print(f"Hauler {device_id}: {result['trips']} trips, {result['trips_per_hour']} trips/hr, {result['total_distance']} km", file=sys.stderr)
            
            elif analyzer.is_excavator(device_id):
                result = analyzer.analyze_excavator_data(device_df)
                if result:
                    excavator_results[str(device_id)] = result
                    print(f"Excavator {device_id}: {result['total_distance']} km", file=sys.stderr)
            
            elif analyzer.is_bulldozer(device_id):
                result = analyzer.analyze_bulldozer_data(device_df)
                if result:
                    bulldozer_results[str(device_id)] = result
                    print(f"Bulldozer {device_id}: {result['total_distance']} km", file=sys.stderr)
        
        if not hauler_results and not excavator_results and not bulldozer_results:
            print(json.dumps({
                'status': 'success',
                'report': '',
                'filename': 'No_Valid_Data.xlsx',
                'devices_discovered': [],
                'record_count': len(df),
                'warning': 'No devices with valid data'
            }))
            return
        
        print("Generating Excel report...", file=sys.stderr)
        generator = ExcelReportGenerator()
        wb = generator.generate_report(hauler_results, excavator_results, bulldozer_results, report_type)
        
        excel_bytes = BytesIO()
        wb.save(excel_bytes)
        excel_bytes.seek(0)
        
        date_str = datetime.now().strftime("%d%b%y").upper()
        type_str = report_type.capitalize()
        filename = f"Mining_Analytics_{type_str}_{date_str}.xlsx"
        
        output = {
            'report': base64.b64encode(excel_bytes.read()).decode('utf-8'),
            'filename': filename,
            'status': 'success',
            'devices_discovered': list(hauler_results.keys()) + list(excavator_results.keys()) + list(bulldozer_results.keys()),
            'record_count': len(df),
            'report_type': report_type,
            'summary': {
                'total_fuel': sum(r.get('total_fuel', 0) for r in hauler_results.values()),
                'total_cost': sum(r.get('total_cost', 0) for r in hauler_results.values()),
                'total_distance': sum(r.get('total_distance', 0) for r in hauler_results.values()),
                'total_trips': sum(r.get('trips', 0) for r in hauler_results.values()),
                'total_hours': sum(r.get('total_hours', r.get('total_shift_hours', 0)) for r in hauler_results.values()),
                'target_met_count': sum(1 for r in hauler_results.values() if r.get('target_met', False))
            }
        }
        
        print(f"✅ Report generated! Haulers: {len(hauler_results)}", file=sys.stderr)
        print(json.dumps(output))
        
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        print(f"ERROR: {str(e)}", file=sys.stderr)
        print(json.dumps({'status': 'error', 'error': str(e)}))

if __name__ == '__main__':
    main()
    '''


    
    
    
    
'''
    # routes/analysis.py - COMPLETE FINAL CODE WITH PROPER SHIFT/DAILY/MONTHLY

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
    def __init__(self):
        self.diesel_price = 94.5
        
        self.fuel_rates = {
            'Steep Up': 1.05,
            'Mild Up': 0.85,
            'Flat': 0.65,
            'Mild Down': 0.45,
            'Steep Down': 0.25
        }
        
        self.max_realistic_speed = 35
        self.max_trips_per_shift = 13  # Maximum 13 trips per 8-hour shift
        
        self.gradient_bands = {
            'Steep Down': (-float('inf'), -8),
            'Mild Down': (-8, -3),
            'Flat': (-3, 3),
            'Mild Up': (3, 8),
            'Steep Up': (8, float('inf'))
        }
        
        # Area definitions
        self.maintenance_lat = 20.4051928
        self.maintenance_lon = 81.0662203
        self.maintenance_radius = 100
        
        self.working_area_lat = 20.4088
        self.working_area_lon = 81.0655
        self.working_area_radius = 150
        
        self.min_trip_distance_km = 0.5
        
        # Equipment classification
        self.hauler_ids = ['D3', 'D10', 'D11', 'D12', '137', '133', '134', '135']
        self.excavator_ids = ['D7', '07', '7', '43']
        self.bulldozer_ids = ['D8', '08']
        
        self.device_mapping = {
            'D3': '137', 'D7': '43', 'D10': '133', 'D11': '134', 'D12': '135'
        }
    
    def map_device_id(self, device_id):
        return self.device_mapping.get(str(device_id).upper(), device_id)
    
    def is_hauler(self, device_id):
        device_str = str(device_id).upper()
        for hid in self.hauler_ids:
            if hid.upper() in device_str:
                return True
        return False
    
    def is_excavator(self, device_id):
        device_str = str(device_id).upper()
        for eid in self.excavator_ids:
            if eid.upper() in device_str:
                return True
        return False
    
    def is_bulldozer(self, device_id):
        device_str = str(device_id).upper()
        for bid in self.bulldozer_ids:
            if bid.upper() in device_str:
                return True
        return False
    
    def is_in_maintenance_area(self, lat, lon):
        if lat == 0 or lon == 0:
            return False
        try:
            dist = geodesic((lat, lon), (self.maintenance_lat, self.maintenance_lon)).meters
            return dist <= self.maintenance_radius
        except:
            return False
    
    def is_in_working_area(self, lat, lon):
        if lat == 0 or lon == 0:
            return False
        try:
            dist = geodesic((lat, lon), (self.working_area_lat, self.working_area_lon)).meters
            return dist <= self.working_area_radius
        except:
            return False
    
    def classify_gradient(self, pitch):
        if pd.isna(pitch):
            return 'Flat'
        if pitch > 30:
            pitch = pitch - 90
        elif pitch < -30:
            pitch = pitch + 90
        pitch = max(-30, min(30, pitch))
        for band_name, (low, high) in self.gradient_bands.items():
            if low < pitch <= high:
                return band_name
        return 'Flat'
    
    def calculate_distance_and_fuel(self, df):
        distances = []
        cumulative_distance = 0
        cumulative_fuel = 0
        
        for i in range(len(df)):
            if i == 0:
                distances.append(0)
                continue
                
            p1 = (df.iloc[i-1]['lat'], df.iloc[i-1]['lon'])
            p2 = (df.iloc[i]['lat'], df.iloc[i]['lon'])
            
            if p1[0] == 0 or p1[1] == 0 or p2[0] == 0 or p2[1] == 0:
                distances.append(0)
                continue
            
            try:
                dist_m = geodesic(p1, p2).meters
                if dist_m > 5000:
                    distances.append(0)
                    continue
                if dist_m < 1:
                    distances.append(0)
                    continue
                    
                distances.append(dist_m)
                dist_km = dist_m / 1000
                cumulative_distance += dist_km
                
                pitch = df.iloc[i]['pitch']
                gradient_class = self.classify_gradient(pitch)
                fuel_rate = self.fuel_rates.get(gradient_class, 0.65)
                segment_fuel = dist_km * fuel_rate
                cumulative_fuel += segment_fuel
                
            except:
                distances.append(0)
        
        df['distance_m'] = distances
        df['distance_km'] = df['distance_m'] / 1000
        df['cumulative_distance_km'] = cumulative_distance
        df['cumulative_fuel_l'] = cumulative_fuel
        df['gradient_class'] = df['pitch'].apply(self.classify_gradient)
        df['fuel_l'] = df.apply(lambda row: row['distance_km'] * self.fuel_rates.get(row['gradient_class'], 0.65), axis=1)
        df['in_maintenance'] = df.apply(lambda row: self.is_in_maintenance_area(row['lat'], row['lon']), axis=1)
        df['in_working_area'] = df.apply(lambda row: self.is_in_working_area(row['lat'], row['lon']), axis=1)
        
        return df
    
    def detect_real_trips(self, df):
        trips = []
        current_trip = 0
        in_trip = False
        trip_distance = 0
        left_working_area = False
        
        for i in range(len(df)):
            lat, lon = df.iloc[i]['lat'], df.iloc[i]['lon']
            in_maintenance = df.iloc[i]['in_maintenance']
            in_working = df.iloc[i]['in_working_area']
            
            if lat == 0 or lon == 0:
                trips.append(current_trip if in_trip else 0)
                continue
            
            if in_maintenance:
                if in_trip:
                    if trip_distance >= self.min_trip_distance_km and left_working_area:
                        trips.append(current_trip)
                    else:
                        trips.append(0)
                        current_trip -= 1
                    in_trip = False
                    trip_distance = 0
                    left_working_area = False
                else:
                    trips.append(0)
                continue
            
            if in_working:
                if not in_trip:
                    trips.append(0)
                elif in_trip and left_working_area:
                    if trip_distance >= self.min_trip_distance_km:
                        trips.append(current_trip)
                    else:
                        trips.append(0)
                        current_trip -= 1
                    in_trip = False
                    trip_distance = 0
                    left_working_area = False
                else:
                    trips.append(current_trip if in_trip else 0)
                continue
            
            if not in_trip:
                current_trip += 1
                in_trip = True
                left_working_area = True
                trips.append(current_trip)
            elif in_trip:
                trip_distance += df.iloc[i]['distance_km']
                trips.append(current_trip)
            else:
                trips.append(0)
        
        trip_counter = 0
        last_trip = 0
        for i in range(len(trips)):
            if trips[i] > last_trip:
                trip_counter += 1
                last_trip = trips[i]
                trips[i] = trip_counter
            elif trips[i] == last_trip:
                trips[i] = trip_counter
            else:
                trips[i] = 0
        
        df['trip'] = trips
        return df
    
    def detect_trips(self, df):
        return self.detect_real_trips(df)
    
    def detect_routes(self, df):
        valid_coords = df[(df['lat'] != 0) & (df['lon'] != 0) & (~df['in_maintenance'])]
        
        if len(valid_coords) < 10:
            return None, None
        
        coords = valid_coords[['lat', 'lon']].values
        
        try:
            eps = 0.0001
            min_samples = 30
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
            return None, None
    
    def analyze_shift_data(self, df, shift_name=""):
        """Analyze a single shift (max 13 trips)"""
        if len(df) < 3:
            return None
        
        try:
            required_cols = ['lat', 'lon', 'pitch', 'speed', 'alt']
            for col in required_cols:
                if col not in df.columns:
                    df[col] = 0
            
            df['pitch'] = pd.to_numeric(df['pitch'], errors='coerce').fillna(0)
            df.loc[df['pitch'] > 30, 'pitch'] = df.loc[df['pitch'] > 30, 'pitch'] - 90
            df.loc[df['pitch'] < -30, 'pitch'] = df.loc[df['pitch'] < -30, 'pitch'] + 90
            
            df = self.calculate_distance_and_fuel(df)
            df = self.detect_trips(df)
            route_a, route_b = self.detect_routes(df)
            
            maintenance_records = df[df['in_maintenance'] == True]
            maintenance_minutes = len(maintenance_records) if 'time' in df.columns else 0
            total_time = len(df)
            maintenance_percentage = (maintenance_minutes / total_time * 100) if total_time > 0 else 0
            
            valid_gps = df[(df['lat'] != 0) & (df['lon'] != 0) & (df['distance_km'] > 0)]
            
            if len(valid_gps) == 0:
                return None
            
            total_distance = valid_gps['distance_km'].sum()
            total_fuel = valid_gps['fuel_l'].sum()
            total_cost = total_fuel * self.diesel_price
            
            # Count valid trips (only those >= 500m) - CAPPED at max_trips_per_shift for SHIFT data
            valid_trips = []
            trip_details = []
            total_trip_distance = 0
            total_trip_fuel = 0
            slow_trips = []
            
            for trip_num in sorted(df[df['trip'] > 0]['trip'].unique()):
                # For shift data, cap at max_trips_per_shift
                if len(valid_trips) >= self.max_trips_per_shift:
                    break
                    
                trip_data = df[df['trip'] == trip_num]
                
                if len(trip_data) < 5:
                    continue
                
                trip_distance = trip_data['distance_km'].sum()
                
                if trip_distance < self.min_trip_distance_km:
                    continue
                
                trip_fuel = trip_data['fuel_l'].sum()
                trip_cost = trip_fuel * self.diesel_price
                
                total_trip_distance += trip_distance
                total_trip_fuel += trip_fuel
                valid_trips.append(trip_num)
                
                trip_duration_hours = 0
                if 'time' in trip_data.columns and len(trip_data) > 1:
                    time_diff = trip_data.iloc[-1]['time'] - trip_data.iloc[0]['time']
                    trip_duration_hours = time_diff.total_seconds() / 3600
                    
                    if trip_duration_hours > 1:
                        slow_trips.append({
                            'trip_num': int(trip_num),
                            'duration_hours': round(trip_duration_hours, 2),
                            'distance_km': round(trip_distance, 2),
                            'fuel_l': round(trip_fuel, 2)
                        })
                
                avg_speed = trip_data['speed'].mean() if 'speed' in trip_data.columns else 0
                if avg_speed > self.max_realistic_speed:
                    avg_speed = self.max_realistic_speed
                
                trip_details.append({
                    'trip_number': int(trip_num),
                    'start_time': str(trip_data.iloc[0]['time']) if 'time' in trip_data.columns else 'N/A',
                    'end_time': str(trip_data.iloc[-1]['time']) if 'time' in trip_data.columns else 'N/A',
                    'duration_hours': round(trip_duration_hours, 2),
                    'duration_minutes': round(trip_duration_hours * 60, 2),
                    'distance_km': round(float(trip_distance), 2),
                    'fuel_l': round(float(trip_fuel), 2),
                    'cost_rs': round(float(trip_cost), 2),
                    'avg_speed': round(float(avg_speed), 1)
                })
            
            unique_trips = len(valid_trips)
            
            # Operating time
            operating_time_hours = 0
            total_time_hours = 0
            
            if 'time' in df.columns and len(df) > 1:
                time_span = df.iloc[-1]['time'] - df.iloc[0]['time']
                total_time_hours = time_span.total_seconds() / 3600
                total_points = len(df)
                if total_points > 0:
                    moving_points = len(df[(df['speed'] > 2) & (df['lat'] != 0) & (~df['in_maintenance'])])
                    operating_time_hours = (moving_points / total_points) * total_time_hours
            
            trips_per_hour = unique_trips / operating_time_hours if operating_time_hours > 0 else 0
            target_met = trips_per_hour >= 1.5
            target_status = "✓ Target Met (≥1.5 trips/op hr)" if target_met else "✗ Target Not Met (<1.5 trips/op hr)"
            
            moving = df[(df['speed'] > 2) & (df['lat'] != 0) & (~df['in_maintenance'])]
            avg_speed = float(moving['speed'].mean()) if len(moving) > 0 else 0
            if avg_speed > self.max_realistic_speed:
                avg_speed = self.max_realistic_speed
            
            fuel_per_km = total_fuel / total_distance if total_distance > 0 else 0
            
            # Gradient analysis
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
            
            idle_points = len(df[(df['speed'] <= 2) & (df['lat'] != 0) & (~df['in_maintenance'])]) if 'speed' in df.columns else 0
            total_points = len(df[(df['lat'] != 0)])
            idle_ratio = idle_points / total_points if total_points > 0 else 0
            
            if idle_ratio > 0.4:
                idle_pattern = "High Idle Time"
            elif idle_ratio > 0.2:
                idle_pattern = "Medium Idle Time"
            else:
                idle_pattern = "Low Idle Time"
            
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
            
            if fuel_per_km < 0.4:
                fuel_efficiency_rating = "Excellent"
            elif fuel_per_km < 0.65:
                fuel_efficiency_rating = "Good"
            elif fuel_per_km < 0.85:
                fuel_efficiency_rating = "Average"
            else:
                fuel_efficiency_rating = "Needs Improvement"
            
            return {
                'total_distance': round(total_distance, 1),
                'total_fuel': round(total_fuel, 1),
                'total_cost': round(total_cost, 2),
                'cost_rs': round(total_cost, 2),
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
                'maintenance_time_minutes': round(maintenance_minutes, 1),
                'maintenance_time_pct': round(maintenance_percentage, 1),
                'in_maintenance': maintenance_percentage > 0,
                'trips_per_hour': round(trips_per_hour, 2),
                'target_status': target_status,
                'target_met': target_met,
                'slow_trips': slow_trips,
                'total_time_hours': round(total_time_hours, 2),
                'operating_time_hours': round(operating_time_hours, 2)
            }
            
        except Exception as e:
            print(f"Error: {e}", file=sys.stderr)
            return None
    
    def analyze_daily_data(self, df):
        """Analyze daily data (multiple shifts, no cap on trips)"""
        if len(df) < 3:
            return None
        
        try:
            required_cols = ['lat', 'lon', 'pitch', 'speed', 'alt']
            for col in required_cols:
                if col not in df.columns:
                    df[col] = 0
            
            df['pitch'] = pd.to_numeric(df['pitch'], errors='coerce').fillna(0)
            df.loc[df['pitch'] > 30, 'pitch'] = df.loc[df['pitch'] > 30, 'pitch'] - 90
            df.loc[df['pitch'] < -30, 'pitch'] = df.loc[df['pitch'] < -30, 'pitch'] + 90
            
            df = self.calculate_distance_and_fuel(df)
            df = self.detect_trips(df)
            route_a, route_b = self.detect_routes(df)
            
            maintenance_records = df[df['in_maintenance'] == True]
            maintenance_minutes = len(maintenance_records) if 'time' in df.columns else 0
            total_time = len(df)
            maintenance_percentage = (maintenance_minutes / total_time * 100) if total_time > 0 else 0
            
            valid_gps = df[(df['lat'] != 0) & (df['lon'] != 0) & (df['distance_km'] > 0)]
            
            if len(valid_gps) == 0:
                return None
            
            total_distance = valid_gps['distance_km'].sum()
            total_fuel = valid_gps['fuel_l'].sum()
            total_cost = total_fuel * self.diesel_price
            
            # Count valid trips - NO CAP for daily data
            valid_trips = []
            trip_details = []
            total_trip_distance = 0
            total_trip_fuel = 0
            slow_trips = []
            
            for trip_num in sorted(df[df['trip'] > 0]['trip'].unique()):
                trip_data = df[df['trip'] == trip_num]
                
                if len(trip_data) < 5:
                    continue
                
                trip_distance = trip_data['distance_km'].sum()
                
                if trip_distance < self.min_trip_distance_km:
                    continue
                
                trip_fuel = trip_data['fuel_l'].sum()
                trip_cost = trip_fuel * self.diesel_price
                
                total_trip_distance += trip_distance
                total_trip_fuel += trip_fuel
                valid_trips.append(trip_num)
                
                trip_duration_hours = 0
                if 'time' in trip_data.columns and len(trip_data) > 1:
                    time_diff = trip_data.iloc[-1]['time'] - trip_data.iloc[0]['time']
                    trip_duration_hours = time_diff.total_seconds() / 3600
                    
                    if trip_duration_hours > 1.5:
                        slow_trips.append({
                            'trip_num': int(trip_num),
                            'duration_hours': round(trip_duration_hours, 2),
                            'distance_km': round(trip_distance, 2),
                            'fuel_l': round(trip_fuel, 2)
                        })
                
                avg_speed = trip_data['speed'].mean() if 'speed' in trip_data.columns else 0
                if avg_speed > self.max_realistic_speed:
                    avg_speed = self.max_realistic_speed
                
                trip_details.append({
                    'trip_number': int(trip_num),
                    'start_time': str(trip_data.iloc[0]['time']) if 'time' in trip_data.columns else 'N/A',
                    'end_time': str(trip_data.iloc[-1]['time']) if 'time' in trip_data.columns else 'N/A',
                    'duration_hours': round(trip_duration_hours, 2),
                    'duration_minutes': round(trip_duration_hours * 60, 2),
                    'distance_km': round(float(trip_distance), 2),
                    'fuel_l': round(float(trip_fuel), 2),
                    'cost_rs': round(float(trip_cost), 2),
                    'avg_speed': round(float(avg_speed), 1)
                })
            
            unique_trips = len(valid_trips)
            
            # Operating time
            operating_time_hours = 0
            total_time_hours = 0
            
            if 'time' in df.columns and len(df) > 1:
                time_span = df.iloc[-1]['time'] - df.iloc[0]['time']
                total_time_hours = time_span.total_seconds() / 3600
                total_points = len(df)
                if total_points > 0:
                    moving_points = len(df[(df['speed'] > 2) & (df['lat'] != 0) & (~df['in_maintenance'])])
                    operating_time_hours = (moving_points / total_points) * total_time_hours
            
            trips_per_hour = unique_trips / operating_time_hours if operating_time_hours > 0 else 0
            target_met = trips_per_hour >= 1.5
            target_status = "✓ Target Met (≥1.5 trips/op hr)" if target_met else "✗ Target Not Met (<1.5 trips/op hr)"
            
            moving = df[(df['speed'] > 2) & (df['lat'] != 0) & (~df['in_maintenance'])]
            avg_speed = float(moving['speed'].mean()) if len(moving) > 0 else 0
            if avg_speed > self.max_realistic_speed:
                avg_speed = self.max_realistic_speed
            
            fuel_per_km = total_fuel / total_distance if total_distance > 0 else 0
            
            # Gradient analysis
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
            
            idle_points = len(df[(df['speed'] <= 2) & (df['lat'] != 0) & (~df['in_maintenance'])]) if 'speed' in df.columns else 0
            total_points = len(df[(df['lat'] != 0)])
            idle_ratio = idle_points / total_points if total_points > 0 else 0
            
            if idle_ratio > 0.4:
                idle_pattern = "High Idle Time"
            elif idle_ratio > 0.2:
                idle_pattern = "Medium Idle Time"
            else:
                idle_pattern = "Low Idle Time"
            
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
            
            if fuel_per_km < 0.4:
                fuel_efficiency_rating = "Excellent"
            elif fuel_per_km < 0.65:
                fuel_efficiency_rating = "Good"
            elif fuel_per_km < 0.85:
                fuel_efficiency_rating = "Average"
            else:
                fuel_efficiency_rating = "Needs Improvement"
            
            return {
                'total_distance': round(total_distance, 1),
                'total_fuel': round(total_fuel, 1),
                'total_cost': round(total_cost, 2),
                'cost_rs': round(total_cost, 2),
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
                'trip_details': trip_details[:50],
                'route_a': route_a,
                'route_b': route_b,
                'maintenance_time_minutes': round(maintenance_minutes, 1),
                'maintenance_time_pct': round(maintenance_percentage, 1),
                'in_maintenance': maintenance_percentage > 0,
                'trips_per_hour': round(trips_per_hour, 2),
                'target_status': target_status,
                'target_met': target_met,
                'slow_trips': slow_trips,
                'total_time_hours': round(total_time_hours, 2),
                'operating_time_hours': round(operating_time_hours, 2)
            }
            
        except Exception as e:
            print(f"Error: {e}", file=sys.stderr)
            return None
    
    def analyze_monthly_data(self, df):
        """Analyze monthly data (aggregated, no cap on trips)"""
        if len(df) < 3:
            return None
        
        try:
            required_cols = ['lat', 'lon', 'pitch', 'speed', 'alt']
            for col in required_cols:
                if col not in df.columns:
                    df[col] = 0
            
            df['pitch'] = pd.to_numeric(df['pitch'], errors='coerce').fillna(0)
            df.loc[df['pitch'] > 30, 'pitch'] = df.loc[df['pitch'] > 30, 'pitch'] - 90
            df.loc[df['pitch'] < -30, 'pitch'] = df.loc[df['pitch'] < -30, 'pitch'] + 90
            
            df = self.calculate_distance_and_fuel(df)
            df = self.detect_trips(df)
            route_a, route_b = self.detect_routes(df)
            
            maintenance_records = df[df['in_maintenance'] == True]
            maintenance_minutes = len(maintenance_records) if 'time' in df.columns else 0
            total_time = len(df)
            maintenance_percentage = (maintenance_minutes / total_time * 100) if total_time > 0 else 0
            
            valid_gps = df[(df['lat'] != 0) & (df['lon'] != 0) & (df['distance_km'] > 0)]
            
            if len(valid_gps) == 0:
                return None
            
            total_distance = valid_gps['distance_km'].sum()
            total_fuel = valid_gps['fuel_l'].sum()
            total_cost = total_fuel * self.diesel_price
            
            # Count valid trips - NO CAP for monthly data
            valid_trips = []
            trip_details = []
            total_trip_distance = 0
            total_trip_fuel = 0
            slow_trips = []
            
            for trip_num in sorted(df[df['trip'] > 0]['trip'].unique()):
                trip_data = df[df['trip'] == trip_num]
                
                if len(trip_data) < 5:
                    continue
                
                trip_distance = trip_data['distance_km'].sum()
                
                if trip_distance < self.min_trip_distance_km:
                    continue
                
                trip_fuel = trip_data['fuel_l'].sum()
                trip_cost = trip_fuel * self.diesel_price
                
                total_trip_distance += trip_distance
                total_trip_fuel += trip_fuel
                valid_trips.append(trip_num)
                
                trip_duration_hours = 0
                if 'time' in trip_data.columns and len(trip_data) > 1:
                    time_diff = trip_data.iloc[-1]['time'] - trip_data.iloc[0]['time']
                    trip_duration_hours = time_diff.total_seconds() / 3600
                    
                    if trip_duration_hours > 2:
                        slow_trips.append({
                            'trip_num': int(trip_num),
                            'duration_hours': round(trip_duration_hours, 2),
                            'distance_km': round(trip_distance, 2),
                            'fuel_l': round(trip_fuel, 2)
                        })
                
                avg_speed = trip_data['speed'].mean() if 'speed' in trip_data.columns else 0
                if avg_speed > self.max_realistic_speed:
                    avg_speed = self.max_realistic_speed
                
                # Only store first 100 trip details for monthly report
                if len(trip_details) < 100:
                    trip_details.append({
                        'trip_number': int(trip_num),
                        'start_time': str(trip_data.iloc[0]['time']) if 'time' in trip_data.columns else 'N/A',
                        'end_time': str(trip_data.iloc[-1]['time']) if 'time' in trip_data.columns else 'N/A',
                        'duration_hours': round(trip_duration_hours, 2),
                        'duration_minutes': round(trip_duration_hours * 60, 2),
                        'distance_km': round(float(trip_distance), 2),
                        'fuel_l': round(float(trip_fuel), 2),
                        'cost_rs': round(float(trip_cost), 2),
                        'avg_speed': round(float(avg_speed), 1)
                    })
            
            unique_trips = len(valid_trips)
            
            # Operating time
            operating_time_hours = 0
            total_time_hours = 0
            
            if 'time' in df.columns and len(df) > 1:
                time_span = df.iloc[-1]['time'] - df.iloc[0]['time']
                total_time_hours = time_span.total_seconds() / 3600
                total_points = len(df)
                if total_points > 0:
                    moving_points = len(df[(df['speed'] > 2) & (df['lat'] != 0) & (~df['in_maintenance'])])
                    operating_time_hours = (moving_points / total_points) * total_time_hours
            
            trips_per_hour = unique_trips / operating_time_hours if operating_time_hours > 0 else 0
            target_met = trips_per_hour >= 1.5
            target_status = "✓ Target Met (≥1.5 trips/op hr)" if target_met else "✗ Target Not Met (<1.5 trips/op hr)"
            
            moving = df[(df['speed'] > 2) & (df['lat'] != 0) & (~df['in_maintenance'])]
            avg_speed = float(moving['speed'].mean()) if len(moving) > 0 else 0
            if avg_speed > self.max_realistic_speed:
                avg_speed = self.max_realistic_speed
            
            fuel_per_km = total_fuel / total_distance if total_distance > 0 else 0
            
            # Gradient analysis
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
            
            idle_points = len(df[(df['speed'] <= 2) & (df['lat'] != 0) & (~df['in_maintenance'])]) if 'speed' in df.columns else 0
            total_points = len(df[(df['lat'] != 0)])
            idle_ratio = idle_points / total_points if total_points > 0 else 0
            
            if idle_ratio > 0.4:
                idle_pattern = "High Idle Time"
            elif idle_ratio > 0.2:
                idle_pattern = "Medium Idle Time"
            else:
                idle_pattern = "Low Idle Time"
            
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
            
            if fuel_per_km < 0.4:
                fuel_efficiency_rating = "Excellent"
            elif fuel_per_km < 0.65:
                fuel_efficiency_rating = "Good"
            elif fuel_per_km < 0.85:
                fuel_efficiency_rating = "Average"
            else:
                fuel_efficiency_rating = "Needs Improvement"
            
            return {
                'total_distance': round(total_distance, 1),
                'total_fuel': round(total_fuel, 1),
                'total_cost': round(total_cost, 2),
                'cost_rs': round(total_cost, 2),
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
                'trip_details': trip_details[:100],
                'route_a': route_a,
                'route_b': route_b,
                'maintenance_time_minutes': round(maintenance_minutes, 1),
                'maintenance_time_pct': round(maintenance_percentage, 1),
                'in_maintenance': maintenance_percentage > 0,
                'trips_per_hour': round(trips_per_hour, 2),
                'target_status': target_status,
                'target_met': target_met,
                'slow_trips': slow_trips,
                'total_time_hours': round(total_time_hours, 2),
                'operating_time_hours': round(operating_time_hours, 2)
            }
            
        except Exception as e:
            print(f"Error: {e}", file=sys.stderr)
            return None
    
    def analyze_excavator_data(self, df):
        if len(df) < 3:
            return None
        
        try:
            df = self.calculate_distance_and_fuel(df)
            
            maintenance_records = df[df['in_maintenance'] == True]
            maintenance_minutes = len(maintenance_records) if 'time' in df.columns else 0
            total_time = len(df)
            maintenance_percentage = (maintenance_minutes / total_time * 100) if total_time > 0 else 0
            
            valid_gps = df[(df['lat'] != 0) & (df['lon'] != 0) & (df['distance_km'] > 0)]
            
            if len(valid_gps) == 0:
                return None
            
            total_distance = valid_gps['distance_km'].sum()
            total_fuel = valid_gps['fuel_l'].sum()
            total_cost = total_fuel * self.diesel_price
            
            idle_points = len(df[(df['speed'] <= 2) & (df['lat'] != 0) & (~df['in_maintenance'])]) if 'speed' in df.columns else 0
            total_points = len(df[(df['lat'] != 0)])
            idle_ratio = idle_points / total_points if total_points > 0 else 0
            
            operating_status = "Stationary Operation" if total_distance < 5 else "Moving Operation"
            
            return {
                'total_distance': round(total_distance, 1),
                'total_fuel': round(total_fuel, 1),
                'total_cost': round(total_cost, 2),
                'fuel_per_km': round(total_fuel / total_distance, 2) if total_distance > 0 else 0,
                'avg_speed': 0,
                'idle_ratio': round(idle_ratio * 100, 1),
                'maintenance_time_pct': round(maintenance_percentage, 1),
                'operating_status': operating_status
            }
            
        except Exception as e:
            return None
    
    def analyze_bulldozer_data(self, df):
        if len(df) < 3:
            return None
        
        try:
            df = self.calculate_distance_and_fuel(df)
            
            total_hours = 0
            operating_hours = 0
            
            if 'time' in df.columns and len(df) > 1:
                time_span = df.iloc[-1]['time'] - df.iloc[0]['time']
                total_hours = time_span.total_seconds() / 3600
                total_points = len(df)
                if total_points > 0:
                    moving_points = len(df[(df['speed'] > 2) & (df['lat'] != 0) & (~df['in_maintenance'])])
                    operating_hours = (moving_points / total_points) * total_hours
            
            valid_gps = df[(df['lat'] != 0) & (df['lon'] != 0) & (df['distance_km'] > 0)]
            
            total_distance = valid_gps['distance_km'].sum() if len(valid_gps) > 0 else 0
            total_fuel = valid_gps['fuel_l'].sum() if len(valid_gps) > 0 else 0
            total_cost = total_fuel * self.diesel_price
            
            fuel_per_hour = total_fuel / operating_hours if operating_hours > 0 else 0
            
            if operating_hours > total_hours * 0.3:
                status = "Active Operation"
            else:
                status = "Standby/Idle"
            
            return {
                'total_distance': round(total_distance, 1),
                'total_fuel': round(total_fuel, 1),
                'total_cost': round(total_cost, 2),
                'operating_hours': round(operating_hours, 2),
                'fuel_per_hour': round(fuel_per_hour, 2),
                'status': status
            }
            
        except Exception as e:
            return None


class ExcelReportGenerator:
    def __init__(self):
        self.wb = Workbook()
        self.setup_styles()
        self.analyzer = SimpleMiningAnalytics()
    
    def setup_styles(self):
        self.header_fill = PatternFill(start_color='2C3E50', end_color='2C3E50', fill_type='solid')
        self.header_font = Font(color='FFFFFF', bold=True, size=11)
        self.centered = Alignment(horizontal='center', vertical='center')
        self.right_aligned = Alignment(horizontal='right', vertical='center')
        self.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        self.good_fill = PatternFill(start_color='2ECC71', end_color='2ECC71', fill_type='solid')
        self.bad_fill = PatternFill(start_color='E74C3C', end_color='E74C3C', fill_type='solid')
        self.warning_fill = PatternFill(start_color='F39C12', end_color='F39C12', fill_type='solid')
        self.maintenance_fill = PatternFill(start_color='3498DB', end_color='3498DB', fill_type='solid')
    
    def create_executive_summary(self, hauler_results):
        ws = self.wb.create_sheet("Executive Summary", 0)
        
        title_cell = ws.cell(row=1, column=1, value="MINING HAULER PERFORMANCE DASHBOARD")
        title_cell.font = Font(size=16, bold=True, color='2C3E50')
        title_cell.alignment = self.centered
        ws.merge_cells('A1:H1')
        
        date_cell = ws.cell(row=2, column=1, value=f"Report Generated: {datetime.now().strftime('%d %B %Y %H:%M')}")
        date_cell.font = Font(size=10, italic=True)
        ws.merge_cells('A2:H2')
        
        maint_note = ws.cell(row=3, column=1, value="📍 TMC Maintenance Area: Lat 20.4051928, Lon 81.0662203")
        maint_note.font = Font(size=9, color='3498DB')
        ws.merge_cells('A3:H3')
        
        working_note = ws.cell(row=4, column=1, value="📍 Working Area (Loading/Dump): Lat 20.4088, Lon 81.0655 (150m radius)")
        working_note.font = Font(size=9, color='27AE60')
        ws.merge_cells('A4:H4')
        
        target_note = ws.cell(row=5, column=1, value="🎯 Target: 1.5 trips per OPERATING hour (Minimum 500m per trip)")
        target_note.font = Font(size=9, bold=True, color='E67E22')
        ws.merge_cells('A5:H5')
        
        total_fuel = sum(r['total_fuel'] for r in hauler_results.values())
        total_cost = sum(r['total_cost'] for r in hauler_results.values())
        total_distance = sum(r['total_distance'] for r in hauler_results.values())
        total_trips = sum(r['trips'] for r in hauler_results.values())
        avg_fuel_per_km = total_fuel / total_distance if total_distance > 0 else 0
        target_met_count = sum(1 for r in hauler_results.values() if r.get('target_met', False))
        
        metrics = [
            ['Total Fuel Consumed', f"{total_fuel:,.1f} Liters", f"₹{total_cost:,.2f}"],
            ['Total Distance Traveled', f"{total_distance:,.1f} km", ''],
            ['Total Trips Completed', f"{total_trips}", ''],
            ['Average Fuel Efficiency', f"{avg_fuel_per_km:.2f} L/km", ''],
            ['Number of Haulers Analyzed', f"{len(hauler_results)}", ''],
            ['', '', ''],
            ['📊 TARGET SUMMARY (1.5 trips per OPERATING hour)', '', ''],
            ['Haulers Meeting Target', f"{target_met_count} / {len(hauler_results)}", f"{target_met_count/len(hauler_results)*100:.0f}%" if hauler_results else '0%'],
        ]
        
        row = 7
        for metric in metrics:
            ws.cell(row=row, column=1, value=metric[0]).font = Font(bold=True)
            ws.cell(row=row, column=2, value=metric[1])
            if len(metric) > 2:
                ws.cell(row=row, column=3, value=metric[2])
            row += 1
        
        for col in range(1, 4):
            ws.column_dimensions[get_column_letter(col)].width = 30
        
        return ws
    
    def create_hauler_performance(self, results):
        ws = self.wb.create_sheet("Hauler Performance")
        
        headers = ['Hauler ID', 'Distance (km)', 'Fuel (L)', 'Cost (₹)', 'Trips', 
                   'Avg Speed (km/h)', 'Operating Hrs', 'Trips/Op Hr', 'Target']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.centered
            cell.border = self.border
        
        row = 2
        for device_id, data in sorted(results.items(), key=lambda x: x[1]['total_fuel'], reverse=True):
            ws.cell(row=row, column=1, value=self.analyzer.map_device_id(device_id))
            ws.cell(row=row, column=2, value=data['total_distance'])
            ws.cell(row=row, column=3, value=data['total_fuel'])
            ws.cell(row=row, column=4, value=data['total_cost'])
            ws.cell(row=row, column=5, value=data['trips'])
            ws.cell(row=row, column=6, value=data['avg_speed'])
            ws.cell(row=row, column=7, value=data.get('operating_time_hours', 0))
            ws.cell(row=row, column=8, value=data.get('trips_per_hour', 0))
            ws.cell(row=row, column=9, value="✓ Yes" if data.get('target_met', False) else "✗ No")
            
            if data.get('target_met', False):
                ws.cell(row=row, column=9).fill = self.good_fill
            else:
                ws.cell(row=row, column=9).fill = self.bad_fill
            
            for col in range(1, 10):
                ws.cell(row=row, column=col).border = self.border
                if col > 1:
                    ws.cell(row=row, column=col).alignment = self.right_aligned
            row += 1
        
        total_trips = sum(v['trips'] for v in results.values())
        total_dist = sum(v['total_distance'] for v in results.values())
        total_fuel = sum(v['total_fuel'] for v in results.values())
        total_cost = sum(v['total_cost'] for v in results.values())
        target_count = sum(1 for v in results.values() if v.get('target_met', False))
        
        row += 1
        ws.cell(row=row, column=1, value="TOTAL / SUMMARY").font = Font(bold=True)
        ws.cell(row=row, column=2, value=round(total_dist, 1))
        ws.cell(row=row, column=3, value=round(total_fuel, 1))
        ws.cell(row=row, column=4, value=round(total_cost, 2))
        ws.cell(row=row, column=5, value=total_trips)
        ws.cell(row=row, column=9, value=f"{target_count}/{len(results)}")
        
        for col in range(1, 10):
            ws.column_dimensions[get_column_letter(col)].width = 14
        
        return ws
    
    def create_excavator_analysis(self, results):
        ws = self.wb.create_sheet("Excavator Analysis")
        
        headers = ['Equipment ID', 'Type', 'Distance (km)', 'Runtime %', 'Maintenance %', 'Operating Status']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.centered
            cell.border = self.border
        
        row = 2
        for device_id, data in results.items():
            ws.cell(row=row, column=1, value=self.analyzer.map_device_id(device_id))
            ws.cell(row=row, column=2, value="Excavator")
            ws.cell(row=row, column=3, value=data['total_distance'])
            ws.cell(row=row, column=4, value=data['idle_ratio'])
            ws.cell(row=row, column=5, value=data.get('maintenance_time_pct', 0))
            ws.cell(row=row, column=6, value=data['operating_status'])
            
            for col in range(1, 10):
                ws.cell(row=row, column=col).border = self.border
                if col > 1:
                    ws.cell(row=row, column=col).alignment = self.right_aligned
            row += 1
        
        if not results:
            ws.cell(row=2, column=1, value="No excavator data available")
        
        for col in range(1, 10):
            ws.column_dimensions[get_column_letter(col)].width = 16
        
        return ws
    
    def create_bulldozer_analysis(self, results):
        ws = self.wb.create_sheet("Bulldozer Analysis")
        
        headers = ['Equipment ID', 'Type', 'Distance (km)', 'Fuel (L)', 'Cost (₹)', 
                   'Operating Hrs', 'Fuel per Hour (L)', 'Status']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.centered
            cell.border = self.border
        
        row = 2
        for device_id, data in results.items():
            ws.cell(row=row, column=1, value=self.analyzer.map_device_id(device_id))
            ws.cell(row=row, column=2, value="Bulldozer")
            ws.cell(row=row, column=3, value=data['total_distance'])
            ws.cell(row=row, column=4, value=data['total_fuel'])
            ws.cell(row=row, column=5, value=data['total_cost'])
            ws.cell(row=row, column=6, value=data['operating_hours'])
            ws.cell(row=row, column=7, value=data['fuel_per_hour'])
            ws.cell(row=row, column=8, value=data['status'])
            
            if data['status'] == "Active Operation":
                ws.cell(row=row, column=8).fill = self.good_fill
            else:
                ws.cell(row=row, column=8).fill = self.warning_fill
            
            for col in range(1, 9):
                ws.cell(row=row, column=col).border = self.border
                if col > 1:
                    ws.cell(row=row, column=col).alignment = self.right_aligned
            row += 1
        
        if not results:
            ws.cell(row=2, column=1, value="No bulldozer data available")
        
        for col in range(1, 9):
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
                
                ws.cell(row=row, column=1, value=self.analyzer.map_device_id(device_id))
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
                ws.cell(row=row, column=1, value=self.analyzer.map_device_id(device_id))
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
                ws.cell(row=row, column=1, value=self.analyzer.map_device_id(comp['device']))
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
            grad_stats = data['gradient_stats']
            steep_up_dist = grad_stats.get('Steep Up', {}).get('distance_km', 0)
            steep_up_fuel = grad_stats.get('Steep Up', {}).get('fuel_l', 0)
            steep_efficiency = steep_up_fuel / steep_up_dist if steep_up_dist > 0 else 0
            
            ws.cell(row=row, column=1, value=self.analyzer.map_device_id(device_id))
            ws.cell(row=row, column=2, value=steep_up_dist)
            ws.cell(row=row, column=3, value=steep_up_fuel)
            ws.cell(row=row, column=4, value=grad_stats.get('Steep Down', {}).get('distance_km', 0))
            ws.cell(row=row, column=5, value=grad_stats.get('Steep Down', {}).get('fuel_l', 0))
            ws.cell(row=row, column=6, value=grad_stats.get('Flat', {}).get('distance_km', 0))
            ws.cell(row=row, column=7, value=grad_stats.get('Flat', {}).get('fuel_l', 0))
            ws.cell(row=row, column=8, value=grad_stats.get('Steep Up', {}).get('percentage', 0))
            ws.cell(row=row, column=9, value=round(steep_efficiency, 2))
            
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
        
        headers = ['Hauler ID', 'Valid Trips', 'Operating Hrs', 'Avg Trip Dist (km)', 
                   'Avg Trip Fuel (L)', 'Avg Trip Time (min)', 'Trips/Op Hr', 'Target']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.centered
            cell.border = self.border
        
        row = 2
        for device_id, data in results.items():
            total_trips = data.get('trips', 0)
            operating_hours = data.get('operating_time_hours', 0)
            total_distance = data.get('total_trip_distance', 0)
            total_fuel = data.get('total_trip_fuel', 0)
            
            avg_distance = total_distance / total_trips if total_trips > 0 else 0
            avg_fuel = total_fuel / total_trips if total_trips > 0 else 0
            
            trip_details = data.get('trip_details', [])
            avg_trip_time = 0
            if trip_details:
                total_duration = sum(t.get('duration_hours', 0) for t in trip_details)
                avg_trip_time = (total_duration / len(trip_details)) * 60 if total_duration > 0 else 0
            
            ws.cell(row=row, column=1, value=self.analyzer.map_device_id(device_id))
            ws.cell(row=row, column=2, value=total_trips)
            ws.cell(row=row, column=3, value=round(operating_hours, 2))
            ws.cell(row=row, column=4, value=round(avg_distance, 2))
            ws.cell(row=row, column=5, value=round(avg_fuel, 2))
            ws.cell(row=row, column=6, value=round(avg_trip_time, 1))
            ws.cell(row=row, column=7, value=data.get('trips_per_hour', 0))
            ws.cell(row=row, column=8, value="✓ Yes" if data.get('target_met', False) else "✗ No")
            
            if data.get('target_met', False):
                ws.cell(row=row, column=8).fill = self.good_fill
            else:
                ws.cell(row=row, column=8).fill = self.bad_fill
            
            if avg_trip_time > 60:
                ws.cell(row=row, column=6).fill = self.bad_fill
            elif avg_trip_time > 45:
                ws.cell(row=row, column=6).fill = self.warning_fill
            elif avg_trip_time > 0:
                ws.cell(row=row, column=6).fill = self.good_fill
            
            for col in range(1, 9):
                ws.cell(row=row, column=col).border = self.border
                if col > 1:
                    ws.cell(row=row, column=col).alignment = self.right_aligned
            row += 1
        
        row += 2
        target_met_count = sum(1 for v in results.values() if v.get('target_met', False))
        ws.merge_cells(f'A{row}:H{row}')
        ws.cell(row=row, column=1, value=f"✅ Haulers Meeting Target (1.5 trips per operating hour): {target_met_count} / {len(results)}")
        ws.cell(row=row, column=1).font = Font(bold=True, size=12)
        
        for col in range(1, 9):
            ws.column_dimensions[get_column_letter(col)].width = 16
        
        return ws
    
    def create_operational_insights(self, results):
        ws = self.wb.create_sheet("Operational Insights")
        
        headers = ['Hauler ID', 'Idle Pattern', 'Idle %', 'Route Type', 'Avg Speed (km/h)', 
                   'Fuel Efficiency', 'Trips/Op Hr', 'Recommendation']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.centered
            cell.border = self.border
        
        row = 2
        for device_id, data in results.items():
            recommendation = []
            if data['fuel_per_km'] > 0.85:
                recommendation.append(f"High fuel: {data['fuel_per_km']} L/km")
            if data['idle_ratio'] > 30:
                recommendation.append(f"Excessive idle: {data['idle_ratio']}%")
            if data['steep_up_percentage'] > 30:
                recommendation.append(f"Steep terrain {data['steep_up_percentage']}%")
            if data['avg_speed'] < 5 and data['trips'] > 0:
                recommendation.append("Low average speed")
            if data.get('maintenance_time_minutes', 0) > 10:
                recommendation.append(f"Maintenance: {data.get('maintenance_time_minutes', 0)} min")
            if not data.get('target_met', False) and data['trips'] > 0:
                need = 1.5 - data.get('trips_per_hour', 0)
                recommendation.append(f"Need +{need:.2f} more trips/op hr")
            
            if data.get('route_a') and data.get('route_b'):
                if data['route_a']['fuel_per_km'] < data['route_b']['fuel_per_km']:
                    recommendation.append("Use Route A")
                else:
                    recommendation.append("Use Route B")
            
            rec_text = "; ".join(recommendation) if recommendation else "Operating normally"
            
            ws.cell(row=row, column=1, value=self.analyzer.map_device_id(device_id))
            ws.cell(row=row, column=2, value=data['idle_pattern'])
            ws.cell(row=row, column=3, value=data['idle_ratio'])
            ws.cell(row=row, column=4, value=data['route_type'])
            ws.cell(row=row, column=5, value=data['avg_speed'])
            ws.cell(row=row, column=6, value=data['fuel_efficiency_rating'])
            ws.cell(row=row, column=7, value=data.get('trips_per_hour', 0))
            ws.cell(row=row, column=8, value=rec_text)
            
            if data.get('target_met', False):
                ws.cell(row=row, column=7).fill = self.good_fill
            else:
                ws.cell(row=row, column=7).fill = self.bad_fill
            
            for col in range(1, 9):
                ws.cell(row=row, column=col).border = self.border
            row += 1
        
        for col in range(1, 9):
            ws.column_dimensions[get_column_letter(col)].width = 16
        ws.column_dimensions[get_column_letter(8)].width = 45
        
        return ws
    
    def generate_report(self, hauler_results, excavator_results, bulldozer_results, report_type="shift"):
        if 'Sheet' in self.wb.sheetnames:
            std = self.wb['Sheet']
            self.wb.remove(std)
        
        if hauler_results:
            self.create_executive_summary(hauler_results)
            self.create_hauler_performance(hauler_results)
            self.create_route_analysis(hauler_results)
            self.create_gradient_analysis(hauler_results)
            self.create_trip_analysis(hauler_results)
            self.create_operational_insights(hauler_results)
            
            # Add report type note
            ws = self.wb["Executive Summary"]
            note_row = ws.max_row + 2
            report_type_text = f"📋 Report Type: {report_type.upper()} Analysis"
            if report_type == "shift":
                report_type_text += " (Max 13 trips per shift)"
            elif report_type == "daily":
                report_type_text += " (Multiple shifts - No trip limit)"
            elif report_type == "monthly":
                report_type_text += " (Aggregated monthly data - No trip limit)"
            
            ws.cell(row=note_row, column=1, value=report_type_text)
            ws.cell(row=note_row, column=1).font = Font(size=10, bold=True, color='2C3E50')
        
        if excavator_results:
            self.create_excavator_analysis(excavator_results)
        
        if bulldozer_results:
            self.create_bulldozer_analysis(bulldozer_results)
        
        return self.wb


def main():
    try:
        input_data = sys.stdin.read()
        if not input_data:
            print(json.dumps({'status': 'error', 'error': 'No input data'}))
            return
        
        params = json.loads(input_data)
        df = pd.DataFrame(params.get('data', []))
        report_type = params.get('report_type', 'shift')  # Get report type from params
        
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
        print(f"Report type: {report_type}", file=sys.stderr)
        
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
        
        hauler_results = {}
        excavator_results = {}
        bulldozer_results = {}
        
        for device_id in df['device_id'].unique():
            if pd.isna(device_id):
                continue
            
            device_df = df[df['device_id'] == device_id].copy()
            
            if len(device_df) < 3:
                print(f"Device {device_id}: insufficient data ({len(device_df)} records)", file=sys.stderr)
                continue
            
            print(f"\n{'='*50}", file=sys.stderr)
            print(f"Analyzing device {device_id} ({len(device_df)} records) with {report_type} analysis", file=sys.stderr)
            
            if analyzer.is_hauler(device_id):
                # Choose analysis method based on report_type
                if report_type == 'shift':
                    result = analyzer.analyze_shift_data(device_df)
                elif report_type == 'daily':
                    result = analyzer.analyze_daily_data(device_df)
                elif report_type == 'monthly':
                    result = analyzer.analyze_monthly_data(device_df)
                else:
                    result = analyzer.analyze_daily_data(device_df)  # Default to daily
                
                if result:
                    result['device_id'] = str(device_id)
                    hauler_results[str(device_id)] = result
                    print(f"  ✓ Hauler {device_id}: {result['trips']} trips, {result['total_distance']} km, {result['total_fuel']} L fuel", file=sys.stderr)
            
            elif analyzer.is_excavator(device_id):
                result = analyzer.analyze_excavator_data(device_df)
                if result:
                    result['device_id'] = str(device_id)
                    excavator_results[str(device_id)] = result
                    print(f"  ✓ Excavator {device_id}: {result['total_distance']} km, {result['total_fuel']} L fuel", file=sys.stderr)
            
            elif analyzer.is_bulldozer(device_id):
                result = analyzer.analyze_bulldozer_data(device_df)
                if result:
                    result['device_id'] = str(device_id)
                    bulldozer_results[str(device_id)] = result
                    print(f"  ✓ Bulldozer {device_id}: {result['total_distance']} km, {result['total_fuel']} L fuel", file=sys.stderr)
        
        if not hauler_results and not excavator_results and not bulldozer_results:
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
        wb = generator.generate_report(hauler_results, excavator_results, bulldozer_results, report_type)
        
        excel_bytes = BytesIO()
        wb.save(excel_bytes)
        excel_bytes.seek(0)
        
        date_str = datetime.now().strftime("%d%b%y").upper()
        type_str = report_type.capitalize()
        filename = f"Mining_Analytics_{type_str}_{date_str}.xlsx"
        
        output = {
            'report': base64.b64encode(excel_bytes.read()).decode('utf-8'),
            'filename': filename,
            'status': 'success',
            'devices_discovered': list(hauler_results.keys()) + list(excavator_results.keys()) + list(bulldozer_results.keys()),
            'record_count': len(df),
            'report_type': report_type,
            'summary': {
                'total_fuel': sum(r['total_fuel'] for r in hauler_results.values()),
                'total_cost': sum(r['total_cost'] for r in hauler_results.values()),
                'total_distance': sum(r['total_distance'] for r in hauler_results.values()),
                'total_trips': sum(r['trips'] for r in hauler_results.values()),
                'target_met_count': sum(1 for r in hauler_results.values() if r.get('target_met', False))
            }
        }
        
        print(f"\n✅ Report generated successfully!", file=sys.stderr)
        print(f"   Haulers: {len(hauler_results)} | Excavators: {len(excavator_results)} | Bulldozers: {len(bulldozer_results)}", file=sys.stderr)
        print(f"   Total Fuel: {output['summary']['total_fuel']:,.2f} L", file=sys.stderr)
        print(f"   Total Cost: ₹{output['summary']['total_cost']:,.2f}", file=sys.stderr)
        print(f"   Total Trips: {output['summary']['total_trips']}", file=sys.stderr)
        
        print(json.dumps(output))
        
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        print(f"ERROR: {str(e)}", file=sys.stderr)
        print(error_trace, file=sys.stderr)
        print(json.dumps({'status': 'error', 'error': str(e), 'traceback': error_trace}))

if __name__ == '__main__':
    main()
    '''
    
    
    
    
    
 
 
 
      
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
        
        # ============================================
        # ADD MAPS BELOW TABLE - WITH PROPER SPACING
        # ============================================
        if trip_image_data:
            map_start_row = row + 2
            
            # Separator
            ws_trip.merge_range(map_start_row, 0, map_start_row, 17, '=' * 80, bold_center)
            map_start_row += 2
            
            # Display each trip with spacing
            for trip_num, image_path, trip_data in trip_image_data:
                # Trip number label only (no extra info)
                ws_trip.merge_range(map_start_row, 0, map_start_row, 17, 
                                   f'TRIP {trip_num}', 
                                   subtitle_format)
                map_start_row += 1
                
                # Insert image with larger size
                ws_trip.insert_image(map_start_row, 0, image_path, {'x_scale': 1.5, 'y_scale': 1.5})
                
                # Add 35 rows of spacing to prevent overlap
                map_start_row += 25
                
                print(f"✅ Map for Hauler {hauler} Trip {trip_num} embedded", file=sys.stderr)
          
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