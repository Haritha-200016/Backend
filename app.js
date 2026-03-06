const express = require('express');
const router = express.Router();
const fetch = require('./routes/fetch');
const insert = require('./routes/insert');
const path = require('path');
const cors = require('cors');
const fs = require('fs');
const db = require('./dao/dao');   
const app = express();

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use("/models", express.static("temp_models"));

app.use(express.json({ limit: "200mb" }));
app.use(express.urlencoded({ extended: true, limit: "200mb" }));


const PORT = 5001;

// Company/region APIs
app.get('/getcompanies', fetch.getcompanies);
app.get('/sectors', fetch.getsectors);
app.get('/getregions', fetch.getregions);

// Dashboard APIs
app.get('/fetch-dashboard-data', insert.fetchDashboardData);
app.get('/api/get_last_10_zaxis', insert.getLast10ZAxis);
app.post('/insert-realtime-data', insert.insertRealtimeData);

// Auth / Registration
app.post('/register', insert.register);
app.post('/signin', insert.signin);
app.post('/forgot-password', insert.forgotPassword);

// New: register FCM token from Flutter
app.post('/register-token', insert.registerToken);

//fetching data based on shift wise, daily, monthly
app.get('/getAll-Devices-MonthlyData',insert.getAllDevicesMonthlyData);
app.get('/getDevice-MonthlyData',insert.getDeviceMonthlyData);
app.get('/getAllDevices-DailyData',insert.getAllDevicesDailyData);
app.get('/getDevice-DailyData',insert.getDeviceDailyData);
app.get('/getAllDevices-ShiftData',insert.getAllDevicesShiftData);
app.get('/getDevice-ShiftData',insert.getDeviceShiftData);



app.get('/generate-analysis', insert.generateAnalysisReport);
// Device list endpoint
app.get('/api/devices', insert.getDevices);
app.get('/fetchDashboardDataby',insert.fetchDashboardDataby)

// Test route

app.post('/test', (req, res) => {
  console.log("Test API called with:", req.body);
  res.json({ message: "Test successful", received: req.body });
});

// Dashboard route (serves dashboard.html with latest data)
// Dashboard route
app.get('/dashboard', (req, res) => {
  // Read the HTML file - NO DATABASE QUERY, NO SCRIPT TAG NEEDED
  let html;
  try {
    html = fs.readFileSync(path.join(__dirname, 'public', 'dashboard.html'), 'utf8');
  } catch (fileErr) {
    console.error('Error reading dashboard.html:', fileErr);
    return res.status(500).send('Failed to load dashboard HTML');
  }
  
  // Send the HTML as-is - the browser will parse URL parameters
  res.send(html);
});
//Root
app.get('/', (req, res) => {
  res.send('');
});

// Start server
app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server is running on http://0.0.0.0:${PORT}`);
  console.log(`📡 24-hour data endpoints available:`);
  console.log(`   GET /fetch-24h-data?device_id=D3&company=TMC&region=Kache`);
  console.log(`   GET /fetch-all-devices-24h?company=TMC&region=Kache`);
});