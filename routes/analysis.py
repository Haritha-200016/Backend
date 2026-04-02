# routes/analysis.py - AI-POWERED VERSION with Excel Output (FIXED)

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
from sklearn.metrics import silhouette_score
from sklearn.linear_model import LinearRegression
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

class AIPoweredMiningAnalytics:
    """AI-Powered Mining Analytics with DBSCAN and Linear Regression"""
    
    def __init__(self):
        self.gradient_bands = {
            'Steep Down': (-float('inf'), -8),
            'Mild Down': (-8, -3),
            'Flat': (-3, 3),
            'Mild Up': (3, 8),
            'Steep Up': (8, float('inf'))
        }
        self.diesel_price = 94.5
        self.start_radius = 50  # meters for trip detection
        self.eps = 0.0003  # DBSCAN sensitivity
        self.gradient_limit = 10  # pitch threshold
        
    def calculate_gps_distance(self, df):
        """Calculate GPS distance between consecutive points"""
        distances = [0]
        
        for i in range(1, len(df)):
            try:
                p1 = (float(df.iloc[i-1]['lat']), float(df.iloc[i-1]['lon']))
                p2 = (float(df.iloc[i]['lat']), float(df.iloc[i]['lon']))
                
                # Skip if coordinates are zero or invalid
                if p1[0] == 0 and p1[1] == 0:
                    distances.append(0)
                    continue
                if p2[0] == 0 and p2[1] == 0:
                    distances.append(0)
                    continue
                
                dist = geodesic(p1, p2).meters
                if dist < 1:
                    distances.append(0)
                else:
                    distances.append(dist)
            except:
                distances.append(0)
                
        df['distance_m'] = distances
        df['distance_km'] = df['distance_m'] / 1000
        return df
    
    def classify_gradient(self, pitch):
        """Classify gradient based on pitch angle"""
        if pd.isna(pitch):
            return 'Flat'
        try:
            for band_name, (low, high) in self.gradient_bands.items():
                if low < pitch <= high:
                    return band_name
        except:
            pass
        return 'Flat'
    
    def detect_start_points_with_dbscan(self, df):
        """Use DBSCAN to automatically detect start/end points"""
        # Filter out zero coordinates
        valid_coords = df[(df['lat'] != 0) & (df['lon'] != 0)]
        if len(valid_coords) < 10:
            # Use first valid point
            first_valid = valid_coords.iloc[0] if len(valid_coords) > 0 else df.iloc[0]
            return (first_valid['lat'], first_valid['lon'])
        
        coords = valid_coords[['lat', 'lon']].values
        
        try:
            # Apply DBSCAN clustering
            db = DBSCAN(eps=self.eps, min_samples=20).fit(coords)
            valid_coords = valid_coords.copy()
            valid_coords['cluster'] = db.labels_
            
            # Filter out noise (-1)
            valid = valid_coords[valid_coords['cluster'] != -1]
            
            if len(valid) > 0:
                # Find main start cluster (most frequent)
                start_cluster = valid['cluster'].value_counts().idxmax()
                start_points = valid[valid['cluster'] == start_cluster]
                
                start_lat = start_points['lat'].mean()
                start_lon = start_points['lon'].mean()
                
                return (start_lat, start_lon)
        except:
            pass
            
        # Fallback: use first valid point
        first_valid = valid_coords.iloc[0] if len(valid_coords) > 0 else df.iloc[0]
        return (first_valid['lat'], first_valid['lon'])
    
    def detect_trips_with_ai(self, df):
        """AI-powered trip detection using DBSCAN and radius"""
        # First, detect start point using DBSCAN
        start_point = self.detect_start_points_with_dbscan(df)
        print(f"  AI Detected Start Point: {start_point}", file=sys.stderr)
        
        # Now detect trips based on distance from start point
        trips = []
        current_trip = 0
        in_trip = False
        
        for i in range(len(df)):
            cur_point = (df.iloc[i]['lat'], df.iloc[i]['lon'])
            # Skip points with zero coordinates
            if cur_point[0] == 0 and cur_point[1] == 0:
                trips.append(current_trip if in_trip else 0)
                continue
                
            dist_from_start = geodesic(cur_point, start_point).meters
            
            # Start a trip when we leave the start area
            if dist_from_start > self.start_radius and not in_trip:
                current_trip += 1
                in_trip = True
                trips.append(current_trip)
            # End trip when we return to start area
            elif dist_from_start < self.start_radius and in_trip and i > 50:
                in_trip = False
                trips.append(current_trip)
            else:
                trips.append(current_trip if in_trip else 0)
        
        df['trip'] = trips
        df['start_point_lat'] = start_point[0]
        df['start_point_lon'] = start_point[1]
        
        return df
    
    def analyze_device(self, device_id, df):
        """Complete AI-powered analysis for a single device"""
        if len(df) < 3:
            return None
        
        try:
            # Calculate GPS distances
            df = self.calculate_gps_distance(df)
            
            # Classify gradients
            df['gradient_class'] = df['pitch'].apply(self.classify_gradient)
            
            # AI-powered trip detection
            df = self.detect_trips_with_ai(df)
            
            # Basic metrics - use only valid GPS points for distance
            valid_gps = df[(df['lat'] != 0) & (df['lon'] != 0)]
            total_distance = float(valid_gps['distance_km'].sum())
            
            # Calculate fuel from cost if fuel column is zero but cost exists
            if 'fuel' in df.columns and df['fuel'].sum() == 0 and 'fuel_cost' in df.columns:
                # Fuel is cost divided by price per liter
                df['fuel'] = df['fuel_cost'] / self.diesel_price
                print(f"  Calculated fuel from cost: {df['fuel'].sum():.2f} L", file=sys.stderr)
            
            total_fuel = float(df['fuel'].sum())
            total_cost = float(df['fuel_cost'].sum()) if 'fuel_cost' in df.columns else total_fuel * self.diesel_price
            
            print(f"  Device {device_id} - Total fuel: {total_fuel:.2f} L, Total cost: ₹{total_cost:.2f}, Total distance: {total_distance:.2f} km", file=sys.stderr)
            
            # Skip if no fuel data
            if total_fuel == 0:
                print(f"  Device {device_id}: No fuel consumption data found, skipping", file=sys.stderr)
                return None
            
            # Trip analysis
            unique_trips = df[df['trip'] > 0]['trip'].nunique()
            trip_details = []
            total_trip_distance = 0
            total_trip_fuel = 0
            
            for trip_num in sorted(df[df['trip'] > 0]['trip'].unique()):
                trip_data = df[df['trip'] == trip_num]
                if len(trip_data) < 5:
                    continue
                
                # Skip trips with zero distance
                trip_distance = trip_data['distance_km'].sum()
                if trip_distance == 0:
                    continue
                
                # Get fuel for this trip
                trip_fuel = trip_data['fuel'].sum()
                trip_cost = trip_data['fuel_cost'].sum() if 'fuel_cost' in trip_data.columns else trip_fuel * self.diesel_price
                
                total_trip_distance += trip_distance
                total_trip_fuel += trip_fuel
                
                grad_stats, flat_stats = self.analyze_gradient_sections(trip_data)
                
                trip_duration = trip_data.iloc[-1]['time'] - trip_data.iloc[0]['time']
                
                trip_details.append({
                    'trip_number': int(trip_num),
                    'start_time': str(trip_data.iloc[0]['time']),
                    'end_time': str(trip_data.iloc[-1]['time']),
                    'duration': str(trip_duration),
                    'duration_hours': round(trip_duration.total_seconds() / 3600, 2),
                    'distance_km': round(float(trip_distance), 2),
                    'fuel_l': round(float(trip_fuel), 2),
                    'cost_rs': round(float(trip_cost), 2),
                    'gradient_distance_km': grad_stats['distance_km'],
                    'flat_distance_km': flat_stats['distance_km'],
                    'avg_speed': round(float(trip_data['speed'].mean()), 1) if 'speed' in trip_data.columns else 0
                })
            
            # Average speed
            avg_speed = 0
            if 'speed' in df.columns:
                moving = df[(df['speed'] > 2) & (df['lat'] != 0) & (df['lon'] != 0)]
                avg_speed = float(moving['speed'].mean()) if len(moving) > 0 else 0
            
            # Fuel efficiency
            fuel_per_km = total_fuel / total_distance if total_distance > 0 else 0
            
            # Gradient analysis
            gradient_stats = {}
            for gradient in self.gradient_bands.keys():
                mask = df['gradient_class'] == gradient
                grad_data = df[mask]
                
                if len(grad_data) > 0:
                    distance = float(grad_data['distance_km'].sum())
                    fuel = float(grad_data['fuel'].sum())
                    cost = float(grad_data['fuel_cost'].sum()) if 'fuel_cost' in grad_data.columns else fuel * self.diesel_price
                    
                    gradient_stats[gradient] = {
                        'distance_km': round(distance, 2),
                        'fuel_l': round(fuel, 2),
                        'cost_rs': round(cost, 2),
                        'fuel_per_km': round(fuel/distance if distance > 0 else 0, 3),
                        'percentage': round(distance/total_distance * 100, 1) if total_distance > 0 else 0
                    }
                else:
                    gradient_stats[gradient] = {
                        'distance_km': 0,
                        'fuel_l': 0,
                        'cost_rs': 0,
                        'fuel_per_km': 0,
                        'percentage': 0
                    }
            
            # Idle analysis
            idle_points = len(df[(df['speed'] <= 2) & (df['lat'] != 0)]) if 'speed' in df.columns else 0
            total_points = len(df[(df['lat'] != 0) & (df['lon'] != 0)])
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
            
            # Performance rating
            if fuel_per_km < 0.6:
                fuel_efficiency_rating = "Excellent"
            elif fuel_per_km < 0.8:
                fuel_efficiency_rating = "Good"
            elif fuel_per_km < 1.0:
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
                'speed_interpretation': 'Speed limited by terrain' if avg_speed < 5 else 'Good operational speed',
                'fuel_efficiency_rating': fuel_efficiency_rating,
                'gradient_stats': gradient_stats,
                'idle_pattern': idle_pattern,
                'idle_ratio': round(idle_ratio * 100, 1),
                'route_type': route_type,
                'route_chars': route_chars,
                'steep_up_percentage': round(steep_up_pct, 1),
                'trip_details': trip_details[:20]
            }
            
        except Exception as e:
            print(f"Error analyzing device {device_id}: {str(e)}", file=sys.stderr)
            import traceback
            traceback.print_exc(file=sys.stderr)
            return None
    
    def analyze_gradient_sections(self, df):
        """Analyze gradient vs flat sections"""
        gradient = df[df['pitch'].abs() > self.gradient_limit]
        flat = df[df['pitch'].abs() <= self.gradient_limit]
        
        grad_stats = {
            'points': len(gradient),
            'distance_km': round(gradient['distance_km'].sum(), 2),
            'fuel_l': round(gradient['fuel'].sum(), 2),
            'avg_pitch': round(gradient['pitch'].mean(), 1) if len(gradient) > 0 else 0
        }
        
        flat_stats = {
            'points': len(flat),
            'distance_km': round(flat['distance_km'].sum(), 2),
            'fuel_l': round(flat['fuel'].sum(), 2),
            'avg_pitch': round(flat['pitch'].mean(), 1) if len(flat) > 0 else 0
        }
        
        return grad_stats, flat_stats

class ExcelReportGenerator:
    """Generate comprehensive Excel report with executive insights"""
    
    def __init__(self):
        self.wb = Workbook()
        self.setup_styles()
        
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
    
    def create_executive_summary(self, results):
        """Create executive summary sheet with key metrics"""
        ws = self.wb.create_sheet("Executive Summary", 0)
        
        # Title
        title_cell = ws.cell(row=1, column=1, value="MINING HAULER PERFORMANCE DASHBOARD")
        title_cell.font = Font(size=16, bold=True, color='2C3E50')
        title_cell.alignment = self.centered
        ws.merge_cells('A1:F1')
        
        # Date
        date_cell = ws.cell(row=2, column=1, value=f"Report Generated: {datetime.now().strftime('%d %B %Y %H:%M')}")
        date_cell.font = Font(size=10, italic=True)
        ws.merge_cells('A2:F2')
        
        # Summary metrics
        total_fuel = sum(r['total_fuel'] for r in results.values())
        total_cost = sum(r['total_cost'] for r in results.values())
        total_distance = sum(r['total_distance'] for r in results.values())
        total_trips = sum(r['trips'] for r in results.values())
        avg_fuel_per_km = total_fuel / total_distance if total_distance > 0 else 0
        
        metrics = [
            ['Total Fuel Consumed', f"{total_fuel:,.1f} Liters", f"₹{total_cost:,.2f}"],
            ['Total Distance Traveled', f"{total_distance:,.1f} km", ''],
            ['Total Trips Completed', f"{total_trips}", ''],
            ['Average Fuel Efficiency', f"{avg_fuel_per_km:.2f} L/km", ''],
            ['Number of Haulers Analyzed', f"{len(results)}", '']
        ]
        
        row = 4
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
        """Create hauler performance comparison sheet"""
        ws = self.wb.create_sheet("Hauler Performance")
        
        headers = ['Hauler ID', 'Distance (km)', 'Fuel (L)', 'Cost (₹)', 'Fuel/km (L/km)', 
                   'Trips', 'Avg Speed (km/h)', 'Efficiency Rating', 'Route Type']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.centered
            cell.border = self.border
        
        row = 2
        for device_id, data in sorted(results.items(), key=lambda x: x[1]['total_fuel'], reverse=True):
            ws.cell(row=row, column=1, value=device_id)
            ws.cell(row=row, column=2, value=data['total_distance'])
            ws.cell(row=row, column=3, value=data['total_fuel'])
            ws.cell(row=row, column=4, value=data['total_cost'])
            ws.cell(row=row, column=5, value=data['fuel_per_km'])
            ws.cell(row=row, column=6, value=data['trips'])
            ws.cell(row=row, column=7, value=data['avg_speed'])
            ws.cell(row=row, column=8, value=data['fuel_efficiency_rating'])
            ws.cell(row=row, column=9, value=data['route_type'])
            
            if data['fuel_efficiency_rating'] == 'Excellent':
                ws.cell(row=row, column=8).fill = self.good_fill
            elif data['fuel_efficiency_rating'] == 'Needs Improvement':
                ws.cell(row=row, column=8).fill = self.bad_fill
            
            for col in range(1, 10):
                ws.cell(row=row, column=col).border = self.border
                if col not in [1, 8, 9]:
                    ws.cell(row=row, column=col).alignment = self.right_aligned
            
            row += 1
        
        for col in range(1, 10):
            ws.column_dimensions[get_column_letter(col)].width = 15
        ws.column_dimensions[get_column_letter(8)].width = 18
        ws.column_dimensions[get_column_letter(9)].width = 18
        
        return ws
    
    def create_fuel_cost_analysis(self, results):
        """Create detailed fuel cost analysis sheet"""
        ws = self.wb.create_sheet("Fuel & Cost Analysis")
        
        headers = ['Hauler ID', 'Fuel (L)', 'Cost (₹)', 'Distance (km)', 'Fuel/km (L/km)',
                   'Trips', 'Avg Trip Fuel (L)', 'Avg Trip Cost (₹)', 'Avg Trip Distance (km)']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.centered
            cell.border = self.border
        
        row = 2
        for device_id, data in sorted(results.items(), key=lambda x: x[1]['total_cost'], reverse=True):
            avg_trip_fuel = data['total_fuel'] / data['trips'] if data['trips'] > 0 else 0
            avg_trip_cost = data['total_cost'] / data['trips'] if data['trips'] > 0 else 0
            avg_trip_distance = data['total_distance'] / data['trips'] if data['trips'] > 0 else 0
            
            ws.cell(row=row, column=1, value=device_id)
            ws.cell(row=row, column=2, value=round(data['total_fuel'], 1))
            ws.cell(row=row, column=3, value=round(data['total_cost'], 2))
            ws.cell(row=row, column=4, value=round(data['total_distance'], 1))
            ws.cell(row=row, column=5, value=round(data['fuel_per_km'], 2))
            ws.cell(row=row, column=6, value=data['trips'])
            ws.cell(row=row, column=7, value=round(avg_trip_fuel, 1))
            ws.cell(row=row, column=8, value=round(avg_trip_cost, 2))
            ws.cell(row=row, column=9, value=round(avg_trip_distance, 1))
            
            for col in range(1, 10):
                ws.cell(row=row, column=col).border = self.border
                if col > 1:
                    ws.cell(row=row, column=col).alignment = self.right_aligned
            row += 1
        
        # Add totals
        total_fuel = sum(r['total_fuel'] for r in results.values())
        total_cost = sum(r['total_cost'] for r in results.values())
        total_distance = sum(r['total_distance'] for r in results.values())
        total_trips = sum(r['trips'] for r in results.values())
        
        ws.cell(row=row, column=1, value="TOTAL").font = Font(bold=True)
        ws.cell(row=row, column=2, value=round(total_fuel, 1))
        ws.cell(row=row, column=3, value=round(total_cost, 2))
        ws.cell(row=row, column=4, value=round(total_distance, 1))
        ws.cell(row=row, column=6, value=total_trips)
        
        for col in range(1, 10):
            ws.cell(row=row, column=col).border = self.border
        
        for col in range(1, 10):
            ws.column_dimensions[get_column_letter(col)].width = 14
        
        return ws
    
    def create_gradient_analysis(self, results):
        """Create gradient analysis sheet"""
        ws = self.wb.create_sheet("Terrain Analysis")
        
        headers = ['Hauler ID', 'Steep Up (km)', 'Steep Up Fuel (L)', 'Mild Up (km)', 
                   'Flat (km)', 'Mild Down (km)', 'Steep Down (km)', 'Steep Grade %']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.centered
            cell.border = self.border
        
        row = 2
        for device_id, data in results.items():
            grad_stats = data['gradient_stats']
            
            ws.cell(row=row, column=1, value=device_id)
            ws.cell(row=row, column=2, value=grad_stats.get('Steep Up', {}).get('distance_km', 0))
            ws.cell(row=row, column=3, value=grad_stats.get('Steep Up', {}).get('fuel_l', 0))
            ws.cell(row=row, column=4, value=grad_stats.get('Mild Up', {}).get('distance_km', 0))
            ws.cell(row=row, column=5, value=grad_stats.get('Flat', {}).get('distance_km', 0))
            ws.cell(row=row, column=6, value=grad_stats.get('Mild Down', {}).get('distance_km', 0))
            ws.cell(row=row, column=7, value=grad_stats.get('Steep Down', {}).get('distance_km', 0))
            ws.cell(row=row, column=8, value=data['steep_up_percentage'])
            
            for col in range(1, 9):
                ws.cell(row=row, column=col).border = self.border
                if col > 1:
                    ws.cell(row=row, column=col).alignment = self.right_aligned
            row += 1
        
        for col in range(1, 9):
            ws.column_dimensions[get_column_letter(col)].width = 14
        
        return ws
    
    def create_trip_analysis(self, results):
        """Create detailed trip analysis sheet"""
        ws = self.wb.create_sheet("Trip Details")
        
        headers = ['Hauler ID', 'Trip #', 'Distance (km)', 'Fuel (L)', 'Cost (₹)', 
                   'Duration (hrs)', 'Avg Speed (km/h)', 'Gradient (km)', 'Flat (km)']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.centered
            cell.border = self.border
        
        row = 2
        for device_id, data in results.items():
            for trip in data.get('trip_details', []):
                ws.cell(row=row, column=1, value=device_id)
                ws.cell(row=row, column=2, value=trip.get('trip_number', 0))
                ws.cell(row=row, column=3, value=trip.get('distance_km', 0))
                ws.cell(row=row, column=4, value=trip.get('fuel_l', 0))
                ws.cell(row=row, column=5, value=trip.get('cost_rs', 0))
                ws.cell(row=row, column=6, value=trip.get('duration_hours', 0))
                ws.cell(row=row, column=7, value=trip.get('avg_speed', 0))
                ws.cell(row=row, column=8, value=trip.get('gradient_distance_km', 0))
                ws.cell(row=row, column=9, value=trip.get('flat_distance_km', 0))
                
                for col in range(1, 10):
                    ws.cell(row=row, column=col).border = self.border
                    if col > 2:
                        ws.cell(row=row, column=col).alignment = self.right_aligned
                row += 1
        
        for col in range(1, 10):
            ws.column_dimensions[get_column_letter(col)].width = 12
        
        return ws
    
    def create_operational_insights(self, results):
        """Create operational insights sheet"""
        ws = self.wb.create_sheet("Operational Insights")
        
        headers = ['Hauler ID', 'Idle Pattern', 'Idle %', 'Route Type', 'Avg Speed (km/h)', 
                   'Fuel Efficiency', 'Recommendation']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.centered
            cell.border = self.border
        
        row = 2
        for device_id, data in results.items():
            recommendation = []
            if data['fuel_per_km'] > 0.9:
                recommendation.append("High fuel consumption - check routes")
            if data['idle_ratio'] > 30:
                recommendation.append("Excessive idle time")
            if data['steep_up_percentage'] > 30:
                recommendation.append("Steep gradients impacting efficiency")
            if data['avg_speed'] < 5:
                recommendation.append("Low average speed")
            
            rec_text = "; ".join(recommendation) if recommendation else "Operating within normal parameters"
            
            ws.cell(row=row, column=1, value=device_id)
            ws.cell(row=row, column=2, value=data['idle_pattern'])
            ws.cell(row=row, column=3, value=data['idle_ratio'])
            ws.cell(row=row, column=4, value=data['route_type'])
            ws.cell(row=row, column=5, value=data['avg_speed'])
            ws.cell(row=row, column=6, value=data['fuel_efficiency_rating'])
            ws.cell(row=row, column=7, value=rec_text)
            
            for col in range(1, 8):
                ws.cell(row=row, column=col).border = self.border
            row += 1
        
        for col in range(1, 8):
            ws.column_dimensions[get_column_letter(col)].width = 18
        ws.column_dimensions[get_column_letter(7)].width = 35
        
        return ws
    
    def generate_report(self, results):
        """Generate complete Excel report with all sheets"""
        
        if 'Sheet' in self.wb.sheetnames:
            std = self.wb['Sheet']
            self.wb.remove(std)
        
        self.create_executive_summary(results)
        self.create_hauler_performance(results)
        self.create_fuel_cost_analysis(results)
        self.create_gradient_analysis(results)
        self.create_trip_analysis(results)
        self.create_operational_insights(results)
        
        return self.wb

def main():
    """Main entry point"""
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
        print(f"Sample data (first 3 rows):", file=sys.stderr)
        print(df[['device_id', 'fuel', 'fuel_cost']].head(3).to_string(), file=sys.stderr)
        
        # Check if we have fuel data
        if 'fuel' not in df.columns and 'fuel_cost' not in df.columns:
            print("ERROR: No fuel or fuel_cost column found!", file=sys.stderr)
            print(json.dumps({'status': 'error', 'error': 'No fuel data available'}))
            return
        
        # If fuel is zero but fuel_cost exists, calculate fuel from cost
        if 'fuel' in df.columns and df['fuel'].sum() == 0 and 'fuel_cost' in df.columns:
            print(f"Fuel column has zeros, calculating fuel from cost at ₹{94.5}/L", file=sys.stderr)
            df['fuel'] = df['fuel_cost'] / 94.5
        
        # Rename columns
        column_mapping = {
            'timestamp': 'time',
            'latitude': 'lat',
            'longitude': 'lon',
            'altitude': 'alt',
            'roll': 'rl'
        }
        df = df.rename(columns=column_mapping)
        
        # Ensure required columns
        required = ['time', 'lat', 'lon', 'pitch', 'fuel', 'speed', 'alt']
        for col in required:
            if col not in df.columns:
                print(f"Warning: Column '{col}' not found, creating with zeros", file=sys.stderr)
                df[col] = 0
        
        # Handle pitch
        if 'pitch' in df.columns:
            df.loc[df['pitch'] > 30, 'pitch'] = df.loc[df['pitch'] > 30, 'pitch'] - 90
            df.loc[df['pitch'] < -30, 'pitch'] = df.loc[df['pitch'] < -30, 'pitch'] + 90
        
        # Add RL
        if 'rl' not in df.columns:
            df['rl'] = df['alt'] + 525.5
        
        # Convert time
        df['time'] = pd.to_datetime(df['time'])
        df = df.sort_values('time').reset_index(drop=True)
        
        print(f"\n{'='*50}", file=sys.stderr)
        print(f"Processing {len(df)} records", file=sys.stderr)
        print(f"Unique devices: {df['device_id'].unique().tolist()}", file=sys.stderr)
        print(f"Total fuel across all devices: {df['fuel'].sum():,.2f} L", file=sys.stderr)
        print(f"Total cost across all devices: ₹{df['fuel_cost'].sum():,.2f}", file=sys.stderr)
        
        analyzer = AIPoweredMiningAnalytics()
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
            print(f"  Total fuel: {device_df['fuel'].sum():,.2f} L", file=sys.stderr)
            print(f"  Total cost: ₹{device_df['fuel_cost'].sum():,.2f}", file=sys.stderr)
            
            result = analyzer.analyze_device(device_id, device_df)
            if result:
                results[str(device_id)] = result
                print(f"  ✓ Device {device_id}: {result['trips']} trips, {result['total_distance']} km, {result['total_fuel']} L fuel, ₹{result['total_cost']}", file=sys.stderr)
        
        if not results:
            print("\n❌ No valid results generated", file=sys.stderr)
            print(json.dumps({
                'status': 'success',
                'report': '',
                'filename': 'No_Valid_Data.xlsx',
                'devices_discovered': [],
                'record_count': len(df),
                'warning': 'No devices with valid fuel consumption data'
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
        filename = f"AI_Mining_Analytics_{date_str}.xlsx"
        
        output = {
            'report': base64.b64encode(excel_bytes.read()).decode('utf-8'),
            'filename': filename,
            'status': 'success',
            'devices_discovered': list(results.keys()),
            'record_count': len(df),
            'summary': {
                'total_fuel': sum(r['total_fuel'] for r in results.values()),
                'total_cost': sum(r['total_cost'] for r in results.values()),
                'total_distance': sum(r['total_distance'] for r in results.values()),
                'total_trips': sum(r['trips'] for r in results.values())
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