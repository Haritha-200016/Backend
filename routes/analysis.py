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
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from math import radians, sin, cos, sqrt, asin, atan2, degrees
from datetime import datetime

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

# ============================================
# EXCAVATOR POLYGON (Hardcoded locations)
# ============================================
EXCAVATOR_POLYGON = {
    "name": "Excavator Area",
    "buffer": 50,
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
    {"id": 2, "lat": 20.4100060, "lon": 81.0692020, "name": "Dump Point 2", "type": "circle", "radius": 100},
    {"id": 3, "lat": 20.4120070, "lon": 81.0610040, "name": "Dump Point 3", "type": "circle", "radius": 100},
    {"id": 4, "lat": 20.4023750, "lon": 81.0627800, "name": "Dump Point 4", "type": "circle", "radius": 150},
    {"id": 5, "lat": 20.4015760, "lon": 81.0623530, "name": "Dump Point 5", "type": "circle", "radius": 200},
    {"id": 6, "lat": 20.4023410, "lon": 81.0627860, "name": "Dump Point 6", "type": "circle", "radius": 100},
    {"id": 1, "lat": 20.4048420, "lon": 81.0642720, "name": "Dump Point 1", "type": "circle", "radius": 100},
    {"id": 7, "lat": 20.410206, "lon":  81.067661, "name": "Dump Point 7", "type": "circle", "radius": 100}
]

# ============================================
# POLYGON-BASED DUMP LOCATION
# ============================================
POLYGON_DUMP = {
    "id": 8,
    "name": "Dump Point 8 (Polygon)",
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
    'd11': '134',
    'd12': '135',
}

HAULER_DEVICES = ['d3', 'd10', 'd11', 'd12']
EXCAVATOR_DEVICE = 'd7'
HAULER_SHEETS = ['133', '134', '135', '137']
DUMP_RADIUS_METERS = 100
EXCAVATOR_BUFFER_METERS = 50

MAINTENANCE_LAT = 20.404977
MAINTENANCE_LON = 81.066210
MAINTENANCE_RADIUS = 52

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

def analyze_trips(records):
    trips_data = {}
    all_point_data = []
    
    
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
        
        # Get distance from the distance column (values are in KM)
        distance = record.get('distance', 0)
        try:
            distance = float(distance) if distance else 0
        except:
            distance = 0
        
        # Add this with other telemetry data
        fuel_consumption = record.get('fuel_consumption', 0)
        # Then parse it
        try:
            fuel_consumption = float(fuel_consumption) if fuel_consumption else 0
        except:
            fuel_consumption = 0
            
        pitch = record.get('pitch', 0)
        roll = record.get('roll', 0)
        vibration = record.get('vibration', 0)
        fuel = record.get('fuel', 0)
        speed = record.get('speed', 0)
        
        try:
            lat = float(lat)
            lon = float(lon)
            pitch = float(pitch) if pitch else 0
            roll = float(roll) if roll else 0
            vibration = float(vibration) if vibration else 0
            fuel = float(fuel) if fuel else 0
            speed = float(speed) if speed else 0
        except Exception as e:
            print(f"⚠️ Error parsing record: {e}", file=sys.stderr)
            continue
        
        if not is_valid_gps(lat, lon):
            continue
        
        loc_type, loc_name = get_location_type(lat, lon)
        
        # Check if this is inside polygon or buffer
        is_inside_polygon = 'inside_polygon' in loc_name.lower()
        is_within_buffer = 'within_buffer' in loc_name.lower()
        
        point_info = {
            'hauler': get_hauler_number(device_id),
            'timestamp': timestamp,
            'lat': lat,
            'lon': lon,
            'location_type': loc_type,
            'location_name': loc_name,
            'distance': distance,  # Already in KM
            'fuel_consumption': fuel_consumption,
            'pitch': pitch,
            'roll': roll,
            'vibration': vibration,
            'fuel': fuel,
            'speed': speed,
            'is_inside_polygon': is_inside_polygon,
            'is_within_buffer': is_within_buffer
        }
        all_point_data.append(point_info)
        
        if device_id not in trips_data:
            trips_data[device_id] = {
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
                'trip_distance': 0,  # Accumulated distance in KM
                'trip_loaded_distance': 0,
                'trip_start_distance': 0,
                'trip_first_excavator_distance': 0,
                'trip_first_dump_distance': 0,
                'trip_last_dump_distance': 0,
                'trip_end_distance': 0,
                'trip_started_with_buffer': False,
                'trip_dump_location': None,
                'first_dump_found': False,
                'trip_accumulated_distance': 0,  # Accumulate hop distances in KM
                'trip_start_fuel': 0,
                'last_fuel_level': 0
            }
        
        hauler = trips_data[device_id]
        
        if hauler['first_record'] is None:
            hauler['first_record'] = timestamp
        hauler['last_record'] = timestamp
        
        # Store distance
        hauler['all_gps_points'].append({
            'lat': lat,
            'lon': lon,
            'time': timestamp,
            'type': loc_type,
            'name': loc_name,
            'distance': distance,
            'is_inside_polygon': is_inside_polygon,
            'is_within_buffer': is_within_buffer
        })
        
        # Location history for time tracking only
        if hauler['current_location'] != loc_name:
            if hauler['current_location'] is not None and hauler['location_start_time'] is not None:
                duration = calculate_duration(hauler['location_start_time'], timestamp)
                if duration > 0:
                    if 'Maintenance' in hauler['current_location']:
                        hauler['total_maintenance_minutes'] += duration
                    elif 'Location' in hauler['current_location'] and 'Excavator' not in hauler['current_location'] and 'Dump' not in hauler['current_location']:
                        hauler['total_unknown_minutes'] += duration
                    elif 'Dump' in hauler['current_location']:
                        hauler['total_dump_minutes'] += duration
                    elif 'Excavator' in hauler['current_location']:
                        hauler['total_excavator_minutes'] += duration
            
            hauler['current_location'] = loc_name
            hauler['location_start_time'] = timestamp
            hauler['location_start_lat'] = lat
            hauler['location_start_lon'] = lon
        
        # ============================================
        # TRIP LOGIC - ACCUMULATE HOP DISTANCES (in KM)
        # ============================================
        
        if loc_type == 'maintenance' and hauler['current_trip'] is not None:
            hauler['trip_visited_maintenance'] = True
            hauler['maintenance_points'].append((lat, lon, timestamp, distance))
        
        if loc_type == 'maintenance' and hauler['at_dump'] and hauler['current_trip'] is not None:
            hauler['trip_ended_in_maintenance'] = True
        
        # START TRIP - Store first excavator point
        if loc_type == 'excavator' and not hauler['at_excavator'] and hauler['current_trip'] is None:
            if is_inside_polygon or is_within_buffer:
                hauler['at_excavator'] = True
                hauler['in_trip'] = True
                hauler['trip_start_time'] = timestamp
                hauler['excavator_start_time'] = timestamp
                hauler['trip_points'] = []
                hauler['maintenance_points'] = []
                hauler['trip_accumulated_distance'] = 0  # Reset accumulated distance in KM
                hauler['trip_ended_in_maintenance'] = False
                hauler['trip_ended_in_unknown'] = False
                hauler['trip_visited_maintenance'] = False
                hauler['last_trip_point'] = None
                hauler['trip_unknown_points'] = []
                hauler['first_dump_found'] = False
                hauler['trip_start_fuel'] = fuel
                # Store first excavator point distance
                hauler['trip_first_excavator_distance'] = 0
                hauler['trip_start_distance'] = 0
                hauler['trip_started_with_buffer'] = not is_inside_polygon
                
                hauler['current_trip'] = {
                    'excavator_arrival': timestamp,
                    'excavator_lat': lat,
                    'excavator_lon': lon,
                    'started_with_buffer': not is_inside_polygon,
                    'used_inside_polygon': is_inside_polygon
                }
                hauler['trip_points'].append((lat, lon, timestamp, distance, is_inside_polygon, is_within_buffer))
                hauler['last_trip_point'] = (lat, lon, timestamp, distance, is_inside_polygon, is_within_buffer)
                
                if is_inside_polygon:
                    print(f"🚚 {device_id}: TRIP START - Inside Excavator Polygon at {timestamp}", file=sys.stderr)
                else:
                    print(f"🚚 {device_id}: TRIP START - Excavator Buffer (50m) at {timestamp}", file=sys.stderr)
        
        # TRACK ALL POINTS DURING TRIP - ACCUMULATE HOP DISTANCES (in KM)
        elif hauler['current_trip'] is not None:
            hauler['trip_points'].append((lat, lon, timestamp, distance, is_inside_polygon, is_within_buffer))
            
            if loc_type == 'unknown':
                hauler['trip_unknown_points'].append((lat, lon, timestamp, distance))
                
            if fuel > 0:
                hauler['last_fuel_level'] = fuel    
            
            # ACCUMULATE hop distance (distance is in KM)
            if distance > 0 and distance <= 1:  # Max 1km jump
                hauler['trip_accumulated_distance'] += distance
            elif distance > 1:
                print(f"⚠️ Outlier distance ignored: {distance:.2f}km at {timestamp}", file=sys.stderr)
            
            hauler['last_trip_point'] = (lat, lon, timestamp, distance, is_inside_polygon, is_within_buffer)
        
        # AT DUMP - Store first dump point distance
        if loc_type == 'dump' and hauler['current_trip'] is not None and not hauler['at_dump']:
            hauler['at_dump'] = True
            hauler['dump_start_time'] = timestamp
            
            # Store first dump point distance (accumulated so far)
            if not hauler['first_dump_found']:
                hauler['trip_first_dump_distance'] = hauler['trip_accumulated_distance']
                hauler['trip_dump_location'] = loc_name
                hauler['first_dump_found'] = True
            
            # Update last dump distance
            hauler['trip_last_dump_distance'] = hauler['trip_accumulated_distance']
            
            hauler['current_trip']['dump_arrival'] = timestamp
            hauler['current_trip']['dump_lat'] = lat
            hauler['current_trip']['dump_lon'] = lon
            hauler['current_trip']['dump_location'] = loc_name
            
            # Calculate lead distance (accumulated so far) - already in KM
            lead_distance_km = hauler['trip_accumulated_distance']
            
            print(f"🚚 {device_id}: At Dump ({loc_name}) at {timestamp}, Lead: {lead_distance_km:.2f} km", file=sys.stderr)
        
        # END TRIP - Returned to Excavator
        elif loc_type == 'excavator' and hauler['at_dump'] and hauler['current_trip'] is not None:
            if is_inside_polygon or is_within_buffer:
                hauler['at_dump'] = False
                hauler['at_excavator'] = False
                hauler['in_trip'] = False
                
                # End distance is accumulated distance (in KM)
                hauler['trip_end_distance'] = hauler['trip_accumulated_distance']
                
                # CALCULATE ALL DISTANCES (already in KM)
                lead_distance_km = hauler['trip_first_dump_distance']
                cycle_distance_km = hauler['trip_accumulated_distance']
                
                # ===== CALCULATE FUEL CONSUMPTION =====
                start_fuel = hauler.get('trip_start_fuel', 0)
                end_fuel = fuel  # Current fuel level at trip end
                if start_fuel > 0 and end_fuel > 0 and start_fuel > end_fuel:
                    fuel_consumption = (start_fuel - end_fuel) / 5
                else:
                    fuel_consumption = 0
                
                trip = {
                    'excavator_arrival': hauler['current_trip']['excavator_arrival'],
                    'dump_arrival': hauler['current_trip']['dump_arrival'],
                    'trip_end': timestamp,
                    'dump_location': hauler['trip_dump_location'],
                    'loaded_distance_km': round(lead_distance_km, 2),
                    'trip_distance_km': round(cycle_distance_km, 2),
                    'ended_in_maintenance': False,
                    'ended_in_unknown': False,
                    'maintenance_distance_km': 0,
                    'visited_maintenance': hauler['trip_visited_maintenance'],
                    'started_with_buffer': hauler['trip_started_with_buffer'],
                    'ended_with_buffer': not is_inside_polygon,
                    'fuel': round(fuel_consumption, 2),
                    'lift': 0,
                    'avg_tonnes': 35
                }
                
                trip['time_at_excavator'] = calculate_duration(trip['excavator_arrival'], trip['dump_arrival'])
                trip['time_at_dump'] = calculate_duration(trip['dump_arrival'], trip['trip_end'])
                trip['total_trip_time'] = calculate_duration(trip['excavator_arrival'], trip['trip_end'])
                
                hauler['total_loaded_distance_km'] += lead_distance_km
                hauler['total_trip_distance_km'] += cycle_distance_km
                hauler['trips'].append(trip)
                
                print(f"🚚 {device_id}: TRIP END - Returned to Excavator. Lead: {lead_distance_km:.2f}km, Cycle: {cycle_distance_km:.2f}km", file=sys.stderr)
                
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
        
        # END TRIP - Maintenance interruption
        elif loc_type == 'maintenance' and hauler['at_dump'] and hauler['current_trip'] is not None:
            hauler['at_dump'] = False
            hauler['at_excavator'] = False
            hauler['in_trip'] = False
            hauler['trip_ended_in_maintenance'] = True
            
            # End distance is accumulated distance (in KM)
            hauler['trip_end_distance'] = hauler['trip_accumulated_distance']
            
            lead_distance_km = hauler['trip_first_dump_distance']
            cycle_distance_km = hauler['trip_accumulated_distance']
            maintenance_distance_km = hauler['trip_accumulated_distance'] - hauler['trip_last_dump_distance']
            
            # ===== CALCULATE FUEL CONSUMPTION =====
            start_fuel = hauler.get('trip_start_fuel', 0)
            end_fuel = fuel  # Current fuel level at trip end
            if start_fuel > 0 and end_fuel > 0 and start_fuel > end_fuel:
                fuel_consumption = (start_fuel - end_fuel) / 5
            else:
               fuel_consumption = 0
            
            trip = {
                'excavator_arrival': hauler['current_trip']['excavator_arrival'],
                'dump_arrival': hauler['current_trip']['dump_arrival'],
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
                'lift': 0,
                'avg_tonnes': 35
            }
            
            trip['time_at_excavator'] = calculate_duration(trip['excavator_arrival'], trip['dump_arrival'])
            trip['time_at_dump'] = calculate_duration(trip['dump_arrival'], trip['trip_end'])
            trip['total_trip_time'] = calculate_duration(trip['excavator_arrival'], trip['trip_end'])
            
            hauler['total_loaded_distance_km'] += lead_distance_km
            hauler['total_trip_distance_km'] += cycle_distance_km
            hauler['total_maintenance_distance_km'] += maintenance_distance_km
            hauler['trips'].append(trip)
            
            print(f"🚚 {device_id}: TRIP END - Maintenance! Lead: {lead_distance_km:.2f}km, Cycle: {cycle_distance_km:.2f}km, Maint: {maintenance_distance_km:.2f}km", file=sys.stderr)
            
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
        
        if hauler['current_trip'] is None and loc_type != 'maintenance' and loc_type != 'excavator' and loc_type != 'dump':
            hauler['in_trip'] = False
    
    # ============================================
    # HANDLE INCOMPLETE TRIPS (Ended in Unknown)
    # ============================================
    for device_id, hauler in trips_data.items():
        if hauler['current_trip'] is not None and hauler['at_dump']:
            hauler['trip_ended_in_unknown'] = True
            
            hauler['trip_end_distance'] = hauler['trip_accumulated_distance']
            
            lead_distance_km = hauler['trip_first_dump_distance']
            cycle_distance_km = hauler['trip_accumulated_distance']
            
            # ===== CALCULATE FUEL CONSUMPTION =====
            start_fuel = hauler.get('trip_start_fuel', 0)
            end_fuel = hauler.get('last_fuel_level', 0)
            if start_fuel > 0 and end_fuel > 0 and start_fuel > end_fuel:
                fuel_consumption = (start_fuel - end_fuel) / 5
            else:
                fuel_consumption = 0
            
            trip = {
                'excavator_arrival': hauler['current_trip']['excavator_arrival'],
                'dump_arrival': hauler['current_trip']['dump_arrival'],
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
                'lift': 0,
                'avg_tonnes': 35
            }
            
            trip['time_at_excavator'] = calculate_duration(trip['excavator_arrival'], trip['dump_arrival'])
            trip['time_at_dump'] = calculate_duration(trip['dump_arrival'], trip['trip_end'])
            trip['total_trip_time'] = calculate_duration(trip['excavator_arrival'], trip['trip_end'])
            
            hauler['total_loaded_distance_km'] += lead_distance_km
            hauler['total_trip_distance_km'] += cycle_distance_km
            hauler['trips'].append(trip)
            
            print(f"🚚 {device_id}: TRIP END - Unknown (incomplete). Lead: {lead_distance_km:.2f}km, Cycle: {cycle_distance_km:.2f}km", file=sys.stderr)
            
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
    
    return trips_data, all_point_data

def create_excel(trips_data, all_point_data):
    wb = Workbook()
    
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    
    header_font = Font(bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center = Alignment(horizontal='center', vertical='center')
    left_align = Alignment(horizontal='left', vertical='center')
    
    # ============================================
    # HAULER SUMMARY SHEET
    # ============================================
    ws_summary = wb.create_sheet(title="Hauler Summary", index=0)
    
    ws_summary['A1'] = "MINING HAULER - SHIFT ANALYSIS"
    ws_summary['A1'].font = Font(size=14, bold=True)
    ws_summary['A1'].alignment = center 
    ws_summary.merge_cells('A1:K1')
    
    shift_start = ""
    shift_end = ""
    for device_id, device_data in trips_data.items():
        if device_data.get('first_record'):
            shift_start = device_data.get('first_record', '')
            shift_end = device_data.get('last_record', '')
            break
    
    # Row 2: Shift information in bold
    ws_summary['A2'] = f"Shift: {shift_start} → {shift_end}"
    ws_summary['A2'].font = Font(bold=True, size=12)
    ws_summary['A2'].alignment = center 
    ws_summary.merge_cells('A2:K2')
    #row8:
    ws_summary['A8'] = "Time in minutes"
    ws_summary['A8'].font = Font(bold=True, size=12) 
    ws_summary.merge_cells('A8:K8')
    #row9
    ws_summary['A9'] = "Distance in KM"
    ws_summary['A9'].font = Font(bold=True, size=12) 
    ws_summary.merge_cells('A9:K9')
    
    '''
    ws_summary['A3'] = f"Excavator Polygon: {len(EXCAVATOR_POLYGON['coordinates'])} points, Buffer: {EXCAVATOR_BUFFER_METERS}m"
    ws_summary.merge_cells('A3:O3')
    
    ws_summary['A4'] = f"Maintenance Area: {MAINTENANCE_LAT}, {MAINTENANCE_LON} (Radius: {MAINTENANCE_RADIUS}m)"
    ws_summary.merge_cells('A4:O4')
    
    ws_summary['A5'] = f"GPS Validation: Max Jump {MAX_GPS_JUMP_METERS}m (1km) - Jumps >1km ignored"
    ws_summary.merge_cells('A5:O5')
    
    ws_summary['A6'] = "Dump Locations:"
    ws_summary.merge_cells('A6:O6')
    row = 7
    for dump in ALL_DUMP_LOCATIONS:
        if dump['type'] == 'circle':
            ws_summary.cell(row=row, column=1, value=f"  {dump['name']} (Circle)")
            ws_summary.cell(row=row, column=2, value=f"Lat: {dump['lat']}, Lon: {dump['lon']}, Radius: {dump['radius']}m")
        else:
            ws_summary.cell(row=row, column=1, value=f"  {dump['name']} (Polygon)")
            ws_summary.cell(row=row, column=2, value=f"Points: {len(dump['coordinates'])}, Buffer: {dump['buffer']}m")
        ws_summary.merge_cells(f'B{row}:O{row}')
        row += 1
    '''
    headers = ['Hauler', 'Trips', 
           'Lead Distance', 'Cycle Distance', 'Total Distance',
           'Avg Cycle Time', 'Running Time', 'Idle Time', 
           'In Maintenance ', 'Fuel(L)', 'AVG Lift', 'Avg L/Hr', 
           'Avg L/Tonne', 'Tonne/Hr']
    
    row = 4
    for col, header in enumerate(headers, 1):
        cell = ws_summary.cell(row=row, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center
        ws_summary.column_dimensions[chr(64 + col)].width = 16
    
    row += 1
    for hauler in HAULER_SHEETS:
        for device_id, device_data in trips_data.items():
            if device_data.get('hauler') == hauler:
                trips = device_data.get('trips', [])
                
                loaded_dist = device_data.get('total_loaded_distance_km', 0)
                trip_dist = device_data.get('total_trip_distance_km', 0)
                maint_dist = device_data.get('total_maintenance_distance_km', 0)
                total_dist = loaded_dist + trip_dist + maint_dist
                
                # Get other metrics (you'll need to calculate or fetch these)
                running_time = sum(trip.get('total_trip_time', 0) for trip in trips)
                avg_cycle_time = running_time / len(trips) if len(trips) > 0 else 0
                idle_time = device_data.get('idle_time', 0)
                in_maintenance = device_data.get('total_maintenance_minutes', 0)
                fuel = device_data.get('fuel_liters', 0)
                avg_lift = device_data.get('avg_lift', 0)
                avg_l_per_hr = device_data.get('avg_l_per_hr', 0)
                avg_l_per_tonne = device_data.get('avg_l_per_tonne', 0)
                tonne_per_hr = device_data.get('tonne_per_hr', 0)
                
                ws_summary.cell(row=row, column=1, value=f"Hauler {hauler}")
                ws_summary.cell(row=row, column=2, value=len(trips))
                ws_summary.cell(row=row, column=3, value=round(loaded_dist, 2))
                ws_summary.cell(row=row, column=4, value=round(trip_dist, 2))
                ws_summary.cell(row=row, column=5, value=round(total_dist, 2))  # This is Lead + Cycle (no maintenance)
                ws_summary.cell(row=row, column=6, value=round(avg_cycle_time, 2))
                ws_summary.cell(row=row, column=7, value=round(running_time, 2))
                ws_summary.cell(row=row, column=8, value=round(idle_time, 2))
                ws_summary.cell(row=row, column=9, value=round(in_maintenance, 2))
                ws_summary.cell(row=row, column=10, value=round(fuel, 2))
                ws_summary.cell(row=row, column=11, value=round(avg_lift, 2))
                ws_summary.cell(row=row, column=12, value=round(avg_l_per_hr, 2))
                ws_summary.cell(row=row, column=13, value=round(avg_l_per_tonne, 2))
                ws_summary.cell(row=row, column=14, value=round(tonne_per_hr, 2))
                
                
                for col in range(1, 15):
                    ws_summary.cell(row=row, column=col).border = border
                    ws_summary.cell(row=row, column=col).alignment = center
                row += 1
    
    # ============================================
    # TRIP SHEETS FOR EACH HAULER
    # ============================================
    for hauler in HAULER_SHEETS:
        for device_id, device_data in trips_data.items():
            if device_data.get('hauler') != hauler:
                continue
            
            trips = device_data.get('trips', [])
            loaded_dist = device_data.get('total_loaded_distance_km', 0)
            trip_dist = device_data.get('total_trip_distance_km', 0)
            maint_dist = device_data.get('total_maintenance_distance_km', 0)
            total_dist = loaded_dist + trip_dist + maint_dist
            
            ws_trip = wb.create_sheet(title=f"Hauler {hauler} Trips")
            
            ws_trip['A1'] = f"Hauler {hauler} - Trip Details"
            ws_trip['A1'].font = Font(size=12, bold=True)
            ws_trip['A1'].alignment = center 
            ws_trip.merge_cells('A1:K1')
            '''
            ws_trip['A3'] = f"Total Trips: {len(trips)}"
            ws_trip['A4'] = f"Lead Distance: {loaded_dist:.2f} km (Excavator → Dump)"
            ws_trip['A5'] = f"Cycle Distance: {trip_dist:.2f} km (Excavator → Dump → Excavator)"
            ws_trip['A6'] = f"Maintenance Distance: {maint_dist:.2f} km"
            ws_trip['A7'] = f"Total Distance: {total_dist:.2f} km"
            ws_trip['A8'] = f"Shift: {device_data.get('first_record', '')} → {device_data.get('last_record', '')}"
            ws_trip.merge_cells('A3:K3')
            ws_trip.merge_cells('A4:K4')
            ws_trip.merge_cells('A5:K5')
            ws_trip.merge_cells('A6:K6')
            ws_trip.merge_cells('A7:K7')
            ws_trip.merge_cells('A8:K8')
            '''
            trip_headers = ['Trip No', 'Dump Location', 'Lead Distance (km)', 'Cycle Distance (km)',
                           'Lead1(min)', 'Lead2(min)', 'Cycle Time(min)', 'Fuel', 'Lift', 'Avg tonnes',
                           'Maintenance (km)', 'Ended in Maintenance', 'Ended in Unknown']
            
            row = 4
            for col, header in enumerate(trip_headers, 1):
                cell = ws_trip.cell(row=row, column=col, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = border
                cell.alignment = center
                ws_trip.column_dimensions[chr(64 + col)].width = 18
            
            row = 5
            for i, trip in enumerate(trips, 1):
                ws_trip.cell(row=row, column=1, value=i)
                ws_trip.cell(row=row, column=2, value=trip.get('dump_location', 'Unknown'))
                ws_trip.cell(row=row, column=3, value=trip.get('loaded_distance_km', 0))
                ws_trip.cell(row=row, column=4, value=trip.get('trip_distance_km', 0))
                ws_trip.cell(row=row, column=5, value=trip.get('time_at_excavator', 0))
                ws_trip.cell(row=row, column=6, value=trip.get('time_at_dump', 0))
                ws_trip.cell(row=row, column=7, value=trip.get('total_trip_time', 0))
                ws_trip.cell(row=row, column=8, value=round(trip.get('fuel', 0), 2))
                ws_trip.cell(row=row, column=9, value=round(trip.get('lift', 0), 2))
                ws_trip.cell(row=row, column=10, value=round(trip.get('avg_tonnes', 35), 2))
                ws_trip.cell(row=row, column=11, value=trip.get('maintenance_distance_km', 0) if trip.get('maintenance_distance_km', 0) > 0 else '')
                
                ended_in_maint = trip.get('ended_in_maintenance', False)
                ws_trip.cell(row=row, column=12, value="YES" if ended_in_maint else "NO")
                if ended_in_maint:
                    ws_trip.cell(row=row, column=12).fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
                    ws_trip.cell(row=row, column=12).font = Font(color='FFFFFF')
                
                ended_in_unknown = trip.get('ended_in_unknown', False)
                ws_trip.cell(row=row, column=13, value="YES" if ended_in_unknown else "NO")
                if ended_in_unknown:
                    ws_trip.cell(row=row, column=13).fill = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')
                
                for col in range(1, 14):
                    ws_trip.cell(row=row, column=col).border = border
                    ws_trip.cell(row=row, column=col).alignment = center
                row += 1
    '''
    # ============================================
    # POINT-BY-POINT ANALYSIS SHEET
    # ============================================
    ws_points = wb.create_sheet(title="Point-by-Point Analysis")
    
    ws_points['A1'] = "POINT-BY-POINT GPS ANALYSIS WITH TELEMETRY"
    ws_points['A1'].font = Font(size=14, bold=True)
    ws_points['A1'].alignment = center 
    ws_points.merge_cells('A1:L1')
    
    ws_points['A3'] = f"Total Points: {len(all_point_data)}"
    ws_points.merge_cells('A3:L3')
    
    ws_points['A4'] = f"Excavator Polygon: {len(EXCAVATOR_POLYGON['coordinates'])} points, Buffer: {EXCAVATOR_BUFFER_METERS}m"
    ws_points.merge_cells('A4:L4')
    
    point_headers = ['Hauler', 'Timestamp', 'Latitude', 'Longitude', 'Location Type', 'Location Name',
                     'Distance (km)', 'Fuel Consumption (L)', 'Pitch (°)', 'Roll (°)', 'Vibration (m/s²)', 'Fuel (%)', 'Speed (km/h)']
    
    row = 6
    for col, header in enumerate(point_headers, 1):
        cell = ws_points.cell(row=row, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center
        ws_points.column_dimensions[chr(64 + col)].width = 18
    
    row = 7
    for point in all_point_data:
        ws_points.cell(row=row, column=1, value=f"Hauler {point['hauler']}")
        ws_points.cell(row=row, column=2, value=point['timestamp'])
        ws_points.cell(row=row, column=3, value=round(point['lat'], 6))
        ws_points.cell(row=row, column=4, value=round(point['lon'], 6))
        ws_points.cell(row=row, column=5, value=point['location_type'])
        
        if point['location_type'] == 'excavator':
            ws_points.cell(row=row, column=5).fill = PatternFill(start_color='2ECC71', end_color='2ECC71', fill_type='solid')
        elif point['location_type'] == 'dump':
            ws_points.cell(row=row, column=5).fill = PatternFill(start_color='F39C12', end_color='F39C12', fill_type='solid')
        elif point['location_type'] == 'maintenance':
            ws_points.cell(row=row, column=5).fill = PatternFill(start_color='3498DB', end_color='3498DB', fill_type='solid')
        else:
            ws_points.cell(row=row, column=5).fill = PatternFill(start_color='E74C3C', end_color='E74C3C', fill_type='solid')
        
        ws_points.cell(row=row, column=6, value=point['location_name'])
        ws_points.cell(row=row, column=7, value=round(point.get('distance', 0), 2))
        ws_points.cell(row=row, column=8, value=round(point.get('fuel_consumption', 0), 2))
        ws_points.cell(row=row, column=9, value=round(point.get('pitch', 0), 1))
        ws_points.cell(row=row, column=10, value=round(point.get('roll', 0), 1))
        ws_points.cell(row=row, column=11, value=round(point.get('vibration', 0), 3))
        ws_points.cell(row=row, column=12, value=round(point.get('fuel', 0), 1))
        ws_points.cell(row=row, column=13, value=round(point.get('speed', 0), 1))
        
        for col in range(1, 14):
            ws_points.cell(row=row, column=col).border = border
            ws_points.cell(row=row, column=col).alignment = center
        row += 1
    
    row += 2
    ws_points.cell(row=row, column=1, value="POINT ANALYSIS SUMMARY")
    ws_points.cell(row=row, column=1).font = Font(bold=True, size=12)
    ws_points.merge_cells(f'A{row}:L{row}')
    row += 1
    
    loc_counts = {}
    for point in all_point_data:
        loc_type = point['location_type']
        loc_counts[loc_type] = loc_counts.get(loc_type, 0) + 1
    
    ws_points.cell(row=row, column=1, value="Location Type Distribution:")
    ws_points.cell(row=row, column=1).font = Font(bold=True)
    row += 1
    
    for loc_type, count in loc_counts.items():
        ws_points.cell(row=row, column=1, value=f"  {loc_type.upper()}:")
        ws_points.cell(row=row, column=2, value=count)
        pct = (count / len(all_point_data)) * 100 if all_point_data else 0
        ws_points.cell(row=row, column=3, value=f"{pct:.1f}%")
        row += 1
    '''
    # ============================================
    # TERRAIN ANALYSIS SHEET
    # ============================================
    ws_terrain = wb.create_sheet(title="Terrain Analysis")

    ws_terrain['A1'] = "TERRAIN ANALYSIS - PITCH & ROLL"
    ws_terrain['A1'].font = Font(size=14, bold=True)
    ws_terrain['A1'].alignment = center
    ws_terrain.merge_cells('A1:G1')

    # Headers
    terrain_headers = ['Hauler', 'Steep UP (km)', 'Steep Down (km)', 'Flat (km)', 'Max Pitch', 'Min Pitch', 'Avg Pitch']

    row = 3
    for col, header in enumerate(terrain_headers, 1):
        cell = ws_terrain.cell(row=row, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center
        ws_terrain.column_dimensions[chr(64 + col)].width = 18

    # Calculate and fill data for each hauler
    row = 4
    for hauler in HAULER_SHEETS:
        # Check if hauler exists in trips_data
        hauler_exists = False
        for device_id, device_data in trips_data.items():
            if device_data.get('hauler') == hauler:
                hauler_exists = True
                break

        if not hauler_exists:
            continue

        # Initialize terrain data for this hauler
        terrain_data = {
            'steep_up': 0,
            'steep_down': 0,
            'flat': 0,
            'max_pitch': -999,
            'min_pitch': 999,
            'pitch_values': []
        }

        # Loop through all points and filter by hauler
        for point in all_point_data:
            if point.get('hauler') != hauler:
                continue

            distance = point.get('distance', 0)
            pitch = point.get('pitch', 0)

            # Skip invalid distances
            if distance <= 0 or distance > 1:
                continue
            # Classify terrain
            if pitch > 8:
                terrain_data['steep_up'] += distance
            elif pitch < -8:
                terrain_data['steep_down'] += distance
            else:
                terrain_data['flat'] += distance

            # Track max/min pitch
            if pitch > terrain_data['max_pitch']:
                terrain_data['max_pitch'] = pitch
            if pitch < terrain_data['min_pitch']:
                terrain_data['min_pitch'] = pitch

            terrain_data['pitch_values'].append(pitch)

        # Calculate average pitch
        avg_pitch = sum(terrain_data['pitch_values']) / len(terrain_data['pitch_values']) if terrain_data['pitch_values'] else 0

        # Fill the row
        ws_terrain.cell(row=row, column=1, value=f"Hauler {hauler}")
        ws_terrain.cell(row=row, column=2, value=round(terrain_data['steep_up'], 2))
        ws_terrain.cell(row=row, column=3, value=round(terrain_data['steep_down'], 2))
        ws_terrain.cell(row=row, column=4, value=round(terrain_data['flat'], 2))
        ws_terrain.cell(row=row, column=5, value=round(terrain_data['max_pitch'], 1))
        ws_terrain.cell(row=row, column=6, value=round(terrain_data['min_pitch'], 1))
        ws_terrain.cell(row=row, column=7, value=round(avg_pitch, 1))

        for col in range(1, 8):
            ws_terrain.cell(row=row, column=col).border = border
            ws_terrain.cell(row=row, column=col).alignment = center
        row += 1

    # Add note at bottom
    note_row = row + 1
    ws_terrain.cell(row=note_row, column=1, value="Note: Terrain classification based on pitch values:")
    ws_terrain.merge_cells(f'A{note_row}:G{note_row}')
    note_row += 1
    ws_terrain.cell(row=note_row, column=1, value="  Steep UP: Pitch > 8°")
    ws_terrain.merge_cells(f'A{note_row}:G{note_row}')
    note_row += 1
    ws_terrain.cell(row=note_row, column=1, value="  Steep Down: Pitch < -8°")
    ws_terrain.merge_cells(f'A{note_row}:G{note_row}')
    note_row += 1
    ws_terrain.cell(row=note_row, column=1, value="  Flat: -8° ≤ Pitch ≤ 8°")
    ws_terrain.merge_cells(f'A{note_row}:G{note_row}')
    
    
    # ============================================
    # TIME BREAKDOWN SHEET
    # ============================================
    ws_breakdown = wb.create_sheet(title="Time Breakdown")
    
    ws_breakdown['A1'] = "TIME BREAKDOWN - 8 HOUR SHIFT"
    ws_breakdown['A1'].font = Font(size=14, bold=True)
    ws_breakdown['A1'].alignment = center 
    ws_breakdown.merge_cells('A1:F1')
    
    headers = ['Hauler', 'At Excavator (min)', 'At Dump (min)', 'In Maintenance (min)', 'Travel/Unknown (min)', 'Total (min)']
    row = 3
    for col, header in enumerate(headers, 1):
        cell = ws_breakdown.cell(row=row, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center
        ws_breakdown.column_dimensions[chr(64 + col)].width = 18
    
    row = 4
    for hauler in HAULER_SHEETS:
        for device_id, device_data in trips_data.items():
            if device_data.get('hauler') == hauler:
                first_rec = device_data.get('first_record', '')
                last_rec = device_data.get('last_record', '')
                total_time = calculate_duration(first_rec, last_rec) if first_rec and last_rec else 0
                
                ws_breakdown.cell(row=row, column=1, value=f"Hauler {hauler}")
                ws_breakdown.cell(row=row, column=2, value=round(device_data.get('total_excavator_minutes', 0), 1))
                ws_breakdown.cell(row=row, column=3, value=round(device_data.get('total_dump_minutes', 0), 1))
                ws_breakdown.cell(row=row, column=4, value=round(device_data.get('total_maintenance_minutes', 0), 1))
                ws_breakdown.cell(row=row, column=5, value=round(device_data.get('total_unknown_minutes', 0), 1))
                ws_breakdown.cell(row=row, column=6, value=round(total_time, 1))
                
                for col in range(1, 7):
                    ws_breakdown.cell(row=row, column=col).border = border
                    ws_breakdown.cell(row=row, column=col).alignment = center
                row += 1
    
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ============================================
# MAIN
# ============================================

def main():
    try:
        input_data = sys.stdin.read()
        
        if not input_data:
            print(json.dumps({'status': 'error', 'error': 'No input data received'}))
            return
        
        data = json.loads(input_data)
        records = data.get('data', [])
        
        trips_data, all_point_data = analyze_trips(records)
        
        print("\n📈 Trip Summary:", file=sys.stderr)
        total_trips = 0
        for device_id, device_data in trips_data.items():
            trip_count = len(device_data.get('trips', []))
            hauler = device_data.get('hauler', device_id)
            total_trips += trip_count
            loaded_dist = device_data.get('total_loaded_distance_km', 0)
            trip_dist = device_data.get('total_trip_distance_km', 0)
            maint_dist = device_data.get('total_maintenance_distance_km', 0)
            total_dist = loaded_dist + trip_dist + maint_dist
            print(f"  Hauler {hauler}: {trip_count} trips, Lead: {loaded_dist:.2f} km, Cycle: {trip_dist:.2f} km, Maint: {maint_dist:.2f} km, Total: {total_dist:.2f} km", file=sys.stderr)
            
            for i, trip in enumerate(device_data.get('trips', []), 1):
                ended_maint = "YES" if trip.get('ended_in_maintenance', False) else "NO"
                ended_unknown = "YES" if trip.get('ended_in_unknown', False) else "NO"
                started_buffer = "YES" if trip.get('started_with_buffer', False) else "NO"
                maint = trip.get('maintenance_distance_km', 0)
                print(f"    Cycle {i}: Lead={trip.get('loaded_distance_km', 0):.2f}km, Cycle={trip.get('trip_distance_km', 0):.2f}km, Maint={maint:.2f}km, Started={started_buffer}, EndedMaint={ended_maint}, EndedUnknown={ended_unknown}", file=sys.stderr)
        
        print(f"\n📊 Total trips detected: {total_trips}", file=sys.stderr)
        print(f"📊 Total GPS points analyzed: {len(all_point_data)}", file=sys.stderr)
        
        excel_buffer = create_excel(trips_data, all_point_data)
        excel_base64 = base64.b64encode(excel_buffer.getvalue()).decode('utf-8')
        
        print(json.dumps({
            'status': 'success',
            'report': excel_base64,
            'filename': 'Trip_Report.xlsx'
        }))
        
        print("✅ Report generated: Trip_Report.xlsx", file=sys.stderr)
        
    except json.JSONDecodeError as e:
        print(json.dumps({'status': 'error', 'error': f'Invalid JSON: {str(e)}'}), file=sys.stderr)
    except Exception as e:
        import traceback
        print(json.dumps({'status': 'error', 'error': str(e)}), file=sys.stderr)
        print(traceback.format_exc(), file=sys.stderr)

if __name__ == "__main__":
    main()