const ExcelJS = require('exceljs');
const pool = require('../dao/dao');
const db = require('../dao/dao');
const nodemailer = require('nodemailer');
const { exec } = require('child_process');


// ==================== API ENDPOINTS ====================

// 1. REGISTER TOKEN (Unchanged)
const registerToken = (req, res) => {
  const { userId, fcmToken, region, company } = req.body;
  if (!userId || !fcmToken || !region || !company) {
    return res.status(400).json({ error: 'userId, fcmToken, region, and company required' });
  }

  db.query(
    'SELECT phone_no FROM users WHERE user_id = ? AND company_name = ?',
    [userId, company],
    (err, userResults) => {
      if (err) {
        console.error('❌ DB error fetching user:', err.sqlMessage || err);
        return res.status(500).json({ error: 'DB error' });
      }
      if (!userResults.length) {
        return res.status(404).json({ error: `User ${userId} not found for ${company}` });
      }
      const phoneNo = userResults[0].phone_no;

      db.query(
        'SELECT region_id FROM regions WHERE region_name = ? AND company_name = ?',
        [region, company],
        (err, regionResults) => {
          if (err) {
            console.error('❌ DB error fetching region:', err.sqlMessage || err);
            return res.status(500).json({ error: 'DB error' });
          }
          if (!regionResults.length) {
            return res.status(404).json({ error: `Region ${region} not found for ${company}` });
          }
          const regionId = regionResults[0].region_id;

          db.query(
            'INSERT INTO user_regions (phone_no, region_id) VALUES (?, ?) ON DUPLICATE KEY UPDATE region_id = ?',
            [phoneNo, regionId, regionId],
            (err) => {
              if (err) {
                console.error('❌ DB error in user_regions:', err.sqlMessage || err);
                return res.status(500).json({ error: 'DB error' });
              }

              db.query(
                'INSERT INTO user_tokens (user_id, fcm_token) VALUES (?, ?) ON DUPLICATE KEY UPDATE fcm_token = ?',
                [userId, fcmToken, fcmToken],
                (err) => {
                  if (err) {
                    console.error('❌ DB error in user_tokens:', err.sqlMessage || err);
                    return res.status(500).json({ error: 'DB error' });
                  }
                  console.log(`✅ FCM token registered for userId: ${userId} (phone: ${phoneNo}) in region: ${region}`);
                  res.json({ success: true });
                }
              );
            }
          );
        }
      );
    }
  );
};


const insertRealtimeData = (req, res) => {

  const {
    device_id,
    equipment_name,
    latitude,
    longitude,
    altitude,
    speed,
    pitch,
    roll,
    movement,
    vibration,
    fuel,
    pressure,
    count1,
    timestamp,
    fuel_consumption
  } = req.body;

  console.log("📡 RAW BODY:", req.body);

  if (!device_id) {
    return res.status(400).json({ error: "device_id required" });
  }

  const FUEL_PRICE_PER_LITER = 90;
  const SEA_LEVEL_RL = 496;

  /* ========= SAFE NUMBER FUNCTION ========= */

  const safeFloat = (v) => {
    const n = parseFloat(v);
    return isNaN(n) ? null : n;
  };

  const lat = safeFloat(latitude);
  const lon = safeFloat(longitude);

  const alt = safeFloat(altitude);
  const spd = safeFloat(speed);
  const pit = safeFloat(pitch);
  const rol = safeFloat(roll);
  const vib = safeFloat(vibration);
  const pres = safeFloat(pressure);

  const hasGPS = lat !== null && lon !== null;

  const hasCount =
    count1 !== undefined &&
    count1 !== null;

  /* ========= TIMESTAMP CONVERSION TO IST ========= */
  const convertToIST = (ts) => {
    let date;

    if (ts) {
      // If timestamp provided, parse it
      date = new Date(ts);
      console.log("🕐 Original timestamp (UTC):", date.toISOString());
    } else {
      // If no timestamp, use current time
      date = new Date();
      console.log("🕐 No timestamp provided, using current UTC:", date.toISOString());
    }

    // Convert UTC to IST by adding 5 hours 30 minutes
    // MySQL expects format: YYYY-MM-DD HH:MM:SS
    const istTime = new Date(date.getTime() + (5.5 * 60 * 60 * 1000));

    // Format for MySQL
    const year = istTime.getFullYear();
    const month = String(istTime.getMonth() + 1).padStart(2, '0');
    const day = String(istTime.getDate()).padStart(2, '0');
    const hours = String(istTime.getHours()).padStart(2, '0');
    const minutes = String(istTime.getMinutes()).padStart(2, '0');
    const seconds = String(istTime.getSeconds()).padStart(2, '0');

    const formattedIST = `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;

    console.log("🕐 Converted IST:", formattedIST);
    return formattedIST;
  };

  // Get IST timestamp
  const istTimestamp = convertToIST(timestamp);

  /* ================= DISTANCE FUNCTION ================= */

  const haversineKm = (p1, p2) => {

    const [lat1, lon1] = p1;
    const [lat2, lon2] = p2;

    const R = 6371;

    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLon = (lon2 - lon1) * Math.PI / 180;

    const a =
      Math.sin(dLat / 2) ** 2 +
      Math.cos(lat1 * Math.PI / 180) *
      Math.cos(lat2 * Math.PI / 180) *
      Math.sin(dLon / 2) ** 2;

    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));

    return R * c;
  };

  /* ================= GET REGION ================= */

  pool.query(
    `SELECT region_id FROM devices WHERE device_id=?`,
    [device_id],
    (err, regionResult) => {

      if (err) {
        console.error(err);
        return res.status(500).json({ error: "region query error" });
      }

      if (regionResult.length === 0) {
        return res.status(404).json({ error: "device not found" });
      }

      const region_id = regionResult[0].region_id;

      /* ===================================================== */
      /* ============ CASE 1 : ONLY COUNT DEVICE ============== */
      /* ===================================================== */

      if (hasCount && !hasGPS) {

        console.log("🔢 Dump counter device");

        const insertQuery = `
          INSERT INTO realtime_sensor_data
          (device_id, timestamp, region_id, count1)
          VALUES (?, ?, ?, ?)
        `;

        const values = [
          device_id,
          istTimestamp,  // CHANGED: Using IST timestamp
          region_id,
          count1
        ];

        pool.query(insertQuery, values, (err, result) => {

          if (err) {
            console.error(err);
            return res.status(500).json({ error: "insert error" });
          }

          console.log(`✅ Count stored: ${count1} at ${istTimestamp}`);

          return res.json({
            status: "success",
            message: "count stored",
            inserted_id: result.insertId
          });

        });

        return;
      }

      /* ===================================================== */
      /* ================= GPS DEVICE ========================= */
      /* ===================================================== */

      pool.query(
        `
        SELECT latitude, longitude, fuel
        FROM realtime_sensor_data
        WHERE device_id=?
        AND latitude IS NOT NULL
        AND longitude IS NOT NULL
        ORDER BY id DESC
        LIMIT 1
        `,
        [device_id],
        (err, prev) => {

          let distance = 0;
          let fuelUsed = 0;
          let fuelCost = 0;
          let fuelValue = null;
          let rl = null;

          if (prev && prev.length > 0 && hasGPS) {

            const prevLat = parseFloat(prev[0].latitude);
            const prevLon = parseFloat(prev[0].longitude);

            distance = haversineKm(
              [prevLat, prevLon],
              [lat, lon]
            );
          }

          /* ================= FUEL ================= */

          /* ================= FUEL ================= */

          const currentFuel = safeFloat(fuel);
          const previousFuel =
            prev && prev.length > 0
              ? safeFloat(prev[0].fuel)
              : null;

          fuelValue = currentFuel;
          fuelUsed = 0;

          // Calculate consumption only if
          // current fuel is valid and lower than previous fuel
          if (
            currentFuel !== null &&
            previousFuel !== null &&
            currentFuel > 0 &&
            currentFuel < previousFuel
          ) {
            fuelUsed = previousFuel - currentFuel;
          }

          // Ignore these cases:currentFuel === 0,currentFuel > previousFuel,currentFuel === previousFuel

          fuelCost = fuelUsed * FUEL_PRICE_PER_LITER;

          /* ================= RL ================= */

          if (alt !== null) {
            rl = (alt + SEA_LEVEL_RL).toFixed(2);
          }

          /* ================= INSERT ================= */

          const insertQuery = `
          INSERT INTO realtime_sensor_data
          (
            device_id,
            equipment_name,
            timestamp,
            latitude,
            longitude,
            altitude,
            speed,
            pitch,
            roll,
            movement,
            vibration,
            pressure,
            distance,
            fuel,
            fuel_cost,
            rl,
            region_id,
            fuel_consumption,
            count1
          )
          VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
          `;

          const values = [

            device_id,
            equipment_name || null,
            istTimestamp,  // CHANGED: Using IST timestamp
            lat,
            lon,
            alt,
            spd,
            pit,
            rol,
            movement || null,
            vib,
            pres,
            distance,
            fuelValue,
            fuelCost,
            rl,
            region_id,
            fuelUsed,
            count1 !== undefined ? count1 : null
          ];

          pool.query(insertQuery, values, (err, result) => {

            if (err) {
              console.error(err);
              return res.status(500).json({ error: "insert error" });
            }

            console.log(`📍 Distance: ${(distance * 1000).toFixed(2)} m`);
            console.log(`⛽ Fuel: ${(fuelUsed * 1000).toFixed(2)} mL`);
            console.log(`🕐 Stored at IST: ${istTimestamp}`);

            res.json({
              status: "success",
              message: "gps data stored",
              inserted_id: result.insertId
            });

          });

        }
      );

    }
  );

};

// Register endpoint (unchanged)
const register = (req, res) => {
  console.log("📥 POST /register called");
  const { name, phone_no, email, password, sector_name, company_name, region_ids } = req.body;

  if (!name || !phone_no || !email || !password || !sector_name || !company_name || !region_ids || !Array.isArray(region_ids) || region_ids.length === 0) {
    return res.status(400).send({ status: "error", message: "All fields are required" });
  }

  db.query(
    "SELECT region_id FROM regions WHERE company_name = ? AND region_id IN (?)",
    [company_name, region_ids],
    (err, results) => {
      if (err) {
        console.error("DB error:", err);
        return res.status(500).send({ status: "error", message: "DB error" });
      }

      const validRegionIds = results.map(r => r.region_id.toString());
      if (validRegionIds.length !== region_ids.length) {
        return res.status(400).send({ status: "error", message: "Invalid region_ids" });
      }

      db.query(
        "INSERT INTO users (name, phone_no, email, password, company_name, sector_name, access) VALUES (?, ?, ?, ?, ?, ?, 'in progress')",
        [name, phone_no, email, password, company_name, sector_name],
        (err, userResult) => {
          if (err) {
            console.error("DB error:", err);
            return res.status(500).send({ status: "error", message: "DB error: " + err.message });
          }

          const user_id = userResult.insertId;
          const regionValues = region_ids.map(id => [phone_no, id]);

          db.query(
            "INSERT INTO user_regions (phone_no, region_id) VALUES ?",
            [regionValues],
            (err) => {
              if (err) {
                console.error("DB error:", err);
                return res.status(500).send({ status: "error", message: "DB error: " + err.message });
              }

              res.status(201).send({
                status: "success",
                user_id,
                name,
                company_name,
                message: "User successfully registered"
              });
            }
          );
        }
      );
    }
  );
};

// Signin endpoint (unchanged)
const signin = (req, res) => {
  const { phone_no, password } = req.body;

  if (!/^\d{4}$/.test(password)) {
    return res.status(401).json({ message: 'Invalid credentials' });
  }

  if (!phone_no || !password)
    return res.status(400).json({ message: 'Please provide valid credentials.' });

  const query = 'SELECT * FROM users WHERE phone_no = ? AND password = ?';
  db.query(query, [phone_no, password], (err, result) => {
    if (err) return res.status(500).json({ message: 'Database error' });

    if (result.length > 0) {
      const user = result[0];

      const regionQuery = `
        SELECT r.region_name
        FROM user_regions ur
        JOIN regions r ON ur.region_id = r.region_id
        WHERE ur.phone_no = ?`;

      db.query(regionQuery, [phone_no], (err, regions) => {
        if (err) return res.status(500).json({ message: 'Error fetching regions' });

        return res.status(200).json({
          status: 'success',
          message: 'Login successful',
          user: {
            user_id: user.user_id,
            name: user.name,
            phone_no: user.phone_no,
            email: user.email,
            sector_name: user.sector_name,
            company_name: user.company_name,
            access: user.access,
            regions: regions.map(r => r.region_name),
          },
        });
      });
    } else {
      return res.status(401).json({ message: 'Invalid credentials' });
    }
  });
};

// Forgot password endpoint (unchanged)
const forgotPassword = (req, res) => {
  const { phone_no, password } = req.body;
  if (!phone_no || !password)
    return res.status(400).json({ message: 'Phone number and password required' });
  if (!/^\d{4}$/.test(password)) {
    return res.status(400).json({ message: 'PIN must be exactly 4 digits' });
  }

  const cleanPhone = phone_no.trim();

  const checkQuery = `SELECT * FROM users WHERE phone_no = ?`;
  db.query(checkQuery, [cleanPhone], (err, result) => {
    if (err) return res.status(500).json({ message: 'Database error' });

    if (result.length === 0)
      return res.status(404).json({ message: 'User not found' });

    const updateQuery = `UPDATE users SET password = ? WHERE phone_no = ?`;
    db.query(updateQuery, [password, cleanPhone], (err, updateResult) => {
      if (err) return res.status(500).json({ message: 'Error updating password' });

      return res.status(200).json({ message: 'Password updated successfully' });
    });
  });
};


// Get last 5 z-axis values endpoint (unchanged)
const getLast10ZAxis = (req, res) => {
  const { company, region } = req.query;

  if (!company || !region) {
    return res.status(400).json({ error: "Company and region required" });
  }

  const query = `
  SELECT device_id, pitch, speed, fuel, fuel_consumption, timestamp
  FROM (
    SELECT rs.device_id, rs.pitch, rs.speed, rs.fuel, rs.fuel_consumption, rs.timestamp,
           ROW_NUMBER() OVER (PARTITION BY rs.device_id ORDER BY rs.timestamp DESC) as rn
    FROM realtime_sensor_data rs
    JOIN devices d ON rs.device_id = d.device_id
    JOIN regions r ON d.region_id = r.region_id
    WHERE r.company_name = ? AND r.region_name = ?
  ) t
  WHERE rn <= 5
  ORDER BY device_id, timestamp DESC
`;

  pool.query(query, [company, region.trim()], (err, rows) => {
    if (err) {
      console.error('Error fetching last values per Hauler:', err);
      return res.status(500).json({ error: "Error fetching data" });
    }

    const haulerData = {};

    rows.forEach(row => {
      const equipment = row.device_id;

      if (!haulerData[equipment]) {
        haulerData[equipment] = [];
      }

      haulerData[equipment].push({
        pitch: Number(row.pitch) || 0,
        speed: Number(row.speed) || 0,
        fuel: Number(row.fuel) || 0,
        fuel_consumption: Number(row.fuel_consumption) || 0,
        timestamp: row.timestamp
      });
    });

    res.json(haulerData);
  });
};


// ===============================
// Receive PLY Function
// ===============================
const multer = require("multer");
const path = require("path");
const fs = require("fs");


// Make sure folder exists
const MODEL_DIR = path.join(__dirname, "../temp_models");

if (!fs.existsSync(MODEL_DIR)) {
  fs.mkdirSync(MODEL_DIR);
}


// Multer storage
const storage = multer.diskStorage({

  destination: (req, file, cb) => {
    cb(null, MODEL_DIR);
  },

  filename: (req, file, cb) => {

    const unique =
      Date.now() + "-" + Math.round(Math.random() * 1000);

    cb(null, unique + "-" + file.originalname);
  }
});

const upload = multer({
  storage: storage,
  limits: { fileSize: 200 * 1024 * 1024 } // 200MB
});


// ==================== SHIFT-WISE (with next day handling for night shift) ====================

// Get shift data for a specific device
const getDeviceShiftData = (req, res) => {
  const { device_id, shift, region_id } = req.query; // Add region_id parameter

  if (!device_id || !shift || !region_id) {
    return res.status(400).json({ error: "device_id, shift and region_id required" });
  }

  const today = new Date().toISOString().split('T')[0];
  const tomorrow = new Date(Date.now() + 86400000).toISOString().split('T')[0];

  const shifts = {
    'morning': { start: '06:00:00', end: '14:00:00' },
    'afternoon': { start: '14:00:00', end: '22:00:00' },
    'night': { start: '22:00:00', end: '06:00:00' }
  };

  if (!shifts[shift]) {
    return res.status(400).json({ error: "Invalid shift" });
  }

  let query;
  let params;

  if (shift === 'night') {
    query = `
      SELECT * FROM realtime_sensor_data 
      WHERE device_id = ? 
      AND region_id = ?
      AND (
        (DATE(timestamp) = ? AND TIME(timestamp) >= '22:00:00')
        OR
        (DATE(timestamp) = ? AND TIME(timestamp) < '06:00:00')
      )
      ORDER BY timestamp ASC
    `;
    params = [device_id, region_id, today, tomorrow];
  } else {
    query = `
      SELECT * FROM realtime_sensor_data 
      WHERE device_id = ? 
      AND region_id = ?
      AND DATE(timestamp) = ?
      AND TIME(timestamp) BETWEEN ? AND ?
      ORDER BY timestamp ASC
    `;
    params = [device_id, region_id, today, shifts[shift].start, shifts[shift].end];
  }

  db.query(query, params, (err, results) => {
    if (err) {
      console.error("Error:", err);
      return res.status(500).json({ error: "Database error" });
    }

    // Filter out null values from each row (optional)
    const filteredResults = results.map(row => {
      const filteredRow = {};
      Object.keys(row).forEach(key => {
        if (row[key] !== null && row[key] !== undefined) {
          filteredRow[key] = row[key];
        }
      });
      return filteredRow;
    });

    res.json({
      status: "success",
      device_id,
      region_id,
      shift,
      date: today,
      total_records: filteredResults.length,
      data: filteredResults
    });
  });
};

// Get shift data for ALL devices in a region
//complet code wotking final maincode 
/*const getAllDevicesShiftData = (req, res) => {
  // DEBUG: Log everything
  console.log('========================================');
  console.log('getAllDevicesShiftData called at:', new Date().toISOString());
  console.log('Full req.query:', JSON.stringify(req.query, null, 2));
  console.log('shift value:', req.query.shift);
  console.log('region_id value:', req.query.region_id);
  console.log('========================================');

  const { shift, region_id } = req.query;

  if (!shift || !region_id) {
    console.log('❌ Missing parameters!');
    return res.status(400).json({ error: "shift and region_id required" });
  }

  const today = new Date().toISOString().split('T')[0];
  const tomorrow = new Date(Date.now() + 86400000).toISOString().split('T')[0];

  console.log('📅 Calculated today:', today);
  console.log('📅 Calculated tomorrow:', tomorrow);

  const shifts = {
    'morning': { start: '06:00:00', end: '14:00:00' },
    'afternoon': { start: '14:00:00', end: '22:00:00' },
    'night': { start: '22:00:00', end: '06:00:00' }
  };

  if (!shifts[shift]) {
    console.log('❌ Invalid shift:', shift);
    return res.status(400).json({ error: "Invalid shift" });
  }

  let query;
  let params;

  if (shift === 'night') {
    query = `
      SELECT * FROM realtime_sensor_data 
      WHERE region_id = ?
      AND (
        (DATE(timestamp) = ? AND TIME(timestamp) >= '22:00:00')
        OR
        (DATE(timestamp) = ? AND TIME(timestamp) < '06:00:00')
      )
      ORDER BY device_id, timestamp ASC
    `;
    params = [region_id, today, tomorrow];
  } else {
    query = `
      SELECT * FROM realtime_sensor_data 
      WHERE region_id = ?
      AND DATE(timestamp) = ?
      AND TIME(timestamp) BETWEEN ? AND ?
      ORDER BY device_id, timestamp ASC
    `;
    params = [region_id, today, shifts[shift].start, shifts[shift].end];
  }

  console.log('🔍 SQL Query:', query);
  console.log('🔍 Params:', params);

  db.query(query, params, (err, results) => {
    if (err) {
      console.error("Error:", err);
      return res.status(500).json({ error: "Database error" });
    }

    console.log(`✅ Found ${results.length} records for ${shift} shift`);

    if (results.length > 0) {
      console.log('📝 First record timestamp:', results[0].timestamp);
    }

    const filteredResults = results.map(row => {
      const filteredRow = {};
      Object.keys(row).forEach(key => {
        if (row[key] !== null && row[key] !== undefined) {
          filteredRow[key] = row[key];
        }
      });
      return filteredRow;
    });

    res.json({
      status: "success",
      region_id,
      shift,
      date: today,
      total_records: filteredResults.length,
      data: filteredResults
    });
  });
};*/

//this is for coustom date only siftwise 
/*const getAllDevicesShiftData = (req, res) => {
  // DEBUG: Log everything
  console.log('========================================');
  console.log('getAllDevicesShiftData called at:', new Date().toISOString());
  console.log('Full req.query:', JSON.stringify(req.query, null, 2));
  console.log('shift value:', req.query.shift);
  console.log('region_id value:', req.query.region_id);
  console.log('date value:', req.query.date);
  console.log('========================================');

  const { shift, region_id } = req.query;

  if (!shift || !region_id) {
    console.log('❌ Missing parameters!');
    return res.status(400).json({ error: "shift and region_id required" });
  }

  // HARDCODED TO JUNE 30, 2026
  const dateStr = '2026-08-2';
  const nextDayStr = '2026-08-3';

  console.log('📅 Target date:', dateStr);
  console.log('📅 Next day:', nextDayStr);

  const shifts = {
    'morning': { start: '06:00:00', end: '14:00:00' },
    'afternoon': { start: '14:00:00', end: '22:00:00' },
    'night': { start: '22:00:00', end: '06:00:00' }
  };

  if (!shifts[shift]) {
    console.log('❌ Invalid shift:', shift);
    return res.status(400).json({ error: "Invalid shift" });
  }

  let query;
  let params;

  if (shift === 'night') {
    query = `
      SELECT * FROM realtime_sensor_data 
      WHERE region_id = ?
      AND (
        (DATE(timestamp) = ? AND TIME(timestamp) >= '22:00:00')
        OR
        (DATE(timestamp) = ? AND TIME(timestamp) < '06:00:00')
      )
      ORDER BY device_id, timestamp ASC
    `;
    params = [region_id, dateStr, nextDayStr];
  } else {
    query = `
      SELECT * FROM realtime_sensor_data 
      WHERE region_id = ?
      AND DATE(timestamp) = ?
      AND TIME(timestamp) BETWEEN ? AND ?
      ORDER BY device_id, timestamp ASC
    `;
    params = [region_id, dateStr, shifts[shift].start, shifts[shift].end];
  }

  console.log('🔍 SQL Query:', query);
  console.log('🔍 Params:', params);

  db.query(query, params, (err, results) => {
    if (err) {
      console.error("Error:", err);
      return res.status(500).json({ error: "Database error" });
    }

    console.log(`✅ Found ${results.length} records for ${shift} shift on ${dateStr}`);

    if (results.length > 0) {
      console.log('📝 First record timestamp:', results[0].timestamp);
    }

    const filteredResults = results.map(row => {
      const filteredRow = {};
      Object.keys(row).forEach(key => {
        if (row[key] !== null && row[key] !== undefined) {
          filteredRow[key] = row[key];
        }
      });
      return filteredRow;
    });

    res.json({
      status: "success",
      region_id,
      shift,
      date: dateStr,
      total_records: filteredResults.length,
      data: filteredResults
    });
  });
};*/


const getAllDevicesShiftData = (req, res) => {
  // DEBUG: Log everything
  console.log('========================================');
  console.log('getAllDevicesShiftData called at:', new Date().toISOString());
  console.log('Full req.query:', JSON.stringify(req.query, null, 2));
  console.log('shift value:', req.query.shift);
  console.log('region value:', req.query.region);
  console.log('company value:', req.query.company);
  console.log('sector value:', req.query.sector);
  console.log('========================================');

  const { shift, region, company, sector } = req.query;

  if (!shift || !region || !company || !sector) {
    console.log('❌ Missing parameters!');
    return res.status(400).json({ 
      error: "shift, region, company, and sector required" 
    });
  }

  // HARDCODED TO AUGUST 2, 2026
  const dateStr = '2026-08-11';
  const nextDayStr = '2026-08-12';

  console.log('📅 Target date:', dateStr);
  console.log('📅 Next day:', nextDayStr);

  const shifts = {
    'morning': { start: '06:00:00', end: '14:00:00' },
    'afternoon': { start: '14:00:00', end: '22:00:00' },
    'night': { start: '22:00:00', end: '06:00:00' }
  };

  if (!shifts[shift]) {
    console.log('❌ Invalid shift:', shift);
    return res.status(400).json({ error: "Invalid shift" });
  }

  // Step 1: Get all devices for this company, sector, and region
  // FIXED: Join companies table to get sector_name
  const getDevicesQuery = `
    SELECT d.device_id, d.region_id 
    FROM devices d
    JOIN regions r ON d.region_id = r.region_id
    JOIN companies c ON r.company_name = c.company_name
    WHERE c.company_name = ? 
    AND c.sector_name = ?
    AND r.region_name = ?
  `;

  console.log('🔍 Getting devices with query:', getDevicesQuery);
  console.log('🔍 Params:', [company, sector, region]);

  db.query(getDevicesQuery, [company, sector, region], (err, devices) => {
    if (err) {
      console.error("Error fetching devices:", err);
      return res.status(500).json({ error: "Database error fetching devices" });
    }

    if (!devices || devices.length === 0) {
      console.log('⚠️ No devices found for this region');
      return res.json({
        status: "success",
        region,
        shift,
        date: dateStr,
        total_devices: 0,
        total_records: 0,
        data: []
      });
    }

    console.log(`📱 Found ${devices.length} devices in region ${region}`);

    // Extract device IDs and region IDs
    const deviceIds = devices.map(d => d.device_id);
    const regionId = devices[0].region_id; // All devices in same region

    // Step 2: Get sensor data for all devices
    let query;
    let params;

    if (shift === 'night') {
      query = `
        SELECT * FROM realtime_sensor_data 
        WHERE region_id = ?
        AND device_id IN (?)
        AND (
          (DATE(timestamp) = ? AND TIME(timestamp) >= '22:00:00')
          OR
          (DATE(timestamp) = ? AND TIME(timestamp) < '06:00:00')
        )
        ORDER BY device_id, timestamp ASC
      `;
      params = [regionId, deviceIds, dateStr, nextDayStr];
    } else {
      query = `
        SELECT * FROM realtime_sensor_data 
        WHERE region_id = ?
        AND device_id IN (?)
        AND DATE(timestamp) = ?
        AND TIME(timestamp) BETWEEN ? AND ?
        ORDER BY device_id, timestamp ASC
      `;
      params = [regionId, deviceIds, dateStr, shifts[shift].start, shifts[shift].end];
    }

    console.log('🔍 SQL Query:', query);
    console.log('🔍 Params:', params);

    db.query(query, params, (err, results) => {
      if (err) {
        console.error("Error:", err);
        return res.status(500).json({ error: "Database error" });
      }

      console.log(`✅ Found ${results.length} records for ${shift} shift on ${dateStr}`);

      if (results.length > 0) {
        console.log('📝 First record timestamp:', results[0].timestamp);
      }

      const filteredResults = results.map(row => {
        const filteredRow = {};
        Object.keys(row).forEach(key => {
          if (row[key] !== null && row[key] !== undefined) {
            filteredRow[key] = row[key];
          }
        });
        return filteredRow;
      });

      res.json({
        status: "success",
        region,
        company,
        sector,
        shift,
        date: dateStr,
        total_devices: devices.length,
        devices: deviceIds,
        total_records: filteredResults.length,
        data: filteredResults
      });
    });
  });
};
// ==================== DAILY (24hr - full day) ====================

const getDeviceDailyData = (req, res) => {
  const { device_id, region_id } = req.query;

  if (!device_id || !region_id) {
    return res.status(400).json({ error: "device_id and region_id required" });
  }

  const today = new Date();
  const startDate = new Date(today.setHours(6, 0, 0, 0));
  const endDate = new Date(today.setDate(today.getDate() + 1));
  endDate.setHours(6, 0, 0, 0);

  const query = `
    SELECT * FROM realtime_sensor_data 
    WHERE device_id = ?
    AND region_id = ?
    AND timestamp >= ? 
    AND timestamp < ?
    ORDER BY timestamp ASC
  `;

  db.query(query, [
    device_id,
    region_id,
    startDate,
    endDate
  ], (err, results) => {
    if (err) {
      console.error("Error:", err);
      return res.status(500).json({ error: "Database error" });
    }

    const filteredResults = results.map(row => {
      const filteredRow = {};
      Object.keys(row).forEach(key => {
        if (row[key] !== null && row[key] !== undefined) {
          filteredRow[key] = row[key];
        }
      });
      return filteredRow;
    });

    res.json({
      status: "success",
      device_id,
      region_id,
      period: "6:00 AM to 6:00 AM",
      total_records: filteredResults.length,
      data: filteredResults
    });
  });
};

/*const getAllDevicesDailyData = (req, res) => {
  const { region_id } = req.query;

  if (!region_id) {
    return res.status(400).json({ error: "region_id required" });
  }

  const today = new Date();
  const startDate = new Date(today);
  startDate.setHours(6, 0, 0, 0);  // Today 6:00 AM

  const endDate = new Date(today);
  endDate.setDate(endDate.getDate() + 1);
  endDate.setHours(6, 0, 0, 0);  // Tomorrow 6:00 AM

  const startDateTime = startDate.toISOString().slice(0, 19).replace('T', ' ');
  const endDateTime = endDate.toISOString().slice(0, 19).replace('T', ' ');

  const query = `
    SELECT * FROM realtime_sensor_data 
    WHERE region_id = ?
    AND timestamp >= ? 
    AND timestamp < ?
    ORDER BY device_id, timestamp ASC
  `;

  db.query(query, [region_id, startDateTime, endDateTime], (err, results) => {
    if (err) {
      console.error("Error:", err);
      return res.status(500).json({ error: "Database error" });
    }

    // Filter out null values from each row
    const filteredResults = results.map(row => {
      const filteredRow = {};
      Object.keys(row).forEach(key => {
        if (row[key] !== null && row[key] !== undefined) {
          filteredRow[key] = row[key];
        }
      });
      return filteredRow;
    });

    // Group data by device_id
    const groupedByDevice = {};
    filteredResults.forEach(record => {
      if (!groupedByDevice[record.device_id]) {
        groupedByDevice[record.device_id] = [];
      }
      groupedByDevice[record.device_id].push(record);
    });

    res.json({
      status: "success",
      region_id,
      period: {
        from: startDateTime,
        to: endDateTime,
        duration: "24 hours"
      },
      total_records: filteredResults.length,
      devices: Object.keys(groupedByDevice).length,
      data: filteredResults,
      grouped_by_device: groupedByDevice
    });
  });
};*/
const getAllDevicesDailyData = (req, res) => {
  console.log('========================================');
  console.log('getAllDevicesDailyData called at:', new Date().toISOString());
  console.log('Full req.query:', JSON.stringify(req.query, null, 2));
  console.log('========================================');

  const { region, company, sector } = req.query;

  if (!region || !company || !sector) {
    console.log('❌ Missing parameters!');
    return res.status(400).json({ 
      error: "region, company, and sector required" 
    });
  }

  // HARDCODED TO AUGUST 2, 2026
  const dateStr = '2026-09-03';
  const nextDayStr = '2026-09-04';

  console.log('📅 Target date:', dateStr);
  console.log('📅 Next day:', nextDayStr);

  // Get devices with company, sector, region
  const getDevicesQuery = `
    SELECT d.device_id, d.region_id 
    FROM devices d
    JOIN regions r ON d.region_id = r.region_id
    JOIN companies c ON r.company_name = c.company_name
    WHERE c.company_name = ? 
    AND c.sector_name = ?
    AND r.region_name = ?
  `;

  db.query(getDevicesQuery, [company, sector, region], (err, devices) => {
    if (err) {
      console.error("Error fetching devices:", err);
      return res.status(500).json({ error: "Database error fetching devices" });
    }

    if (!devices || devices.length === 0) {
      console.log('⚠️ No devices found for this region');
      return res.json({
        status: "success",
        region,
        date: dateStr,
        total_devices: 0,
        total_records: 0,
        data: []
      });
    }

    console.log(`📱 Found ${devices.length} devices in region ${region}`);

    const deviceIds = devices.map(d => d.device_id);
    const regionId = devices[0].region_id;

    // Get daily data (6 AM to next day 6 AM)
    const query = `
      SELECT * FROM realtime_sensor_data 
      WHERE region_id = ?
      AND device_id IN (?)
      AND timestamp >= ? 
      AND timestamp < ?
      ORDER BY device_id, timestamp ASC
    `;

    const startDateTime = `${dateStr} 06:00:00`;
    const endDateTime = `${nextDayStr} 06:00:00`;

    const params = [regionId, deviceIds, startDateTime, endDateTime];

    console.log('🔍 SQL Query:', query);
    console.log('🔍 Params:', params);

    db.query(query, params, (err, results) => {
      if (err) {
        console.error("Error:", err);
        return res.status(500).json({ error: "Database error" });
      }

      console.log(`✅ Found ${results.length} records for daily data`);

      const filteredResults = results.map(row => {
        const filteredRow = {};
        Object.keys(row).forEach(key => {
          if (row[key] !== null && row[key] !== undefined) {
            filteredRow[key] = row[key];
          }
        });
        return filteredRow;
      });

      res.json({
        status: "success",
        region,
        company,
        sector,
        period: {
          from: startDateTime,
          to: endDateTime,
          duration: "24 hours"
        },
        total_devices: devices.length,
        devices: deviceIds,
        total_records: filteredResults.length,
        data: filteredResults
      });
    });
  });
};
// ==================== MONTHLY (1st to last day OR 1st to today) ====================

// Get monthly data for a specific device
const getDeviceMonthlyData = (req, res) => {
  const { device_id, region_id } = req.query;

  if (!device_id || !region_id) {
    return res.status(400).json({ error: "device_id and region_id required" });
  }

  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;

  // First day of month
  const firstDay = `${year}-${String(month).padStart(2, '0')}-01`;

  // Last day of month
  const lastDay = new Date(year, month, 0).toISOString().split('T')[0];

  const query = `
    SELECT * FROM realtime_sensor_data 
    WHERE device_id = ?
    AND region_id = ?
    AND DATE(timestamp) BETWEEN ? AND ?
    ORDER BY timestamp ASC
  `;

  db.query(query, [device_id, region_id, firstDay, lastDay], (err, results) => {
    if (err) {
      console.error("Error:", err);
      return res.status(500).json({ error: "Database error" });
    }

    // Filter out null values from each row
    const filteredResults = results.map(row => {
      const filteredRow = {};
      Object.keys(row).forEach(key => {
        if (row[key] !== null && row[key] !== undefined) {
          filteredRow[key] = row[key];
        }
      });
      return filteredRow;
    });

    res.json({
      status: "success",
      device_id,
      region_id,
      month: `${year}-${String(month).padStart(2, '0')}`,
      date_range: {
        from: firstDay,
        to: lastDay
      },
      total_records: filteredResults.length,
      data: filteredResults
    });
  });
};


// Get monthly data for ALL devices in a region
const getAllDevicesMonthlyData = (req, res) => {
  const { region_id } = req.query;

  if (!region_id) {
    return res.status(400).json({ error: "region_id required" });
  }

  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;

  const firstDay = `${year}-${String(month).padStart(2, '0')}-01`;
  const lastDay = new Date(year, month, 0).toISOString().split('T')[0];

  const query = `
    SELECT * FROM realtime_sensor_data 
    WHERE region_id = ?
    AND DATE(timestamp) BETWEEN ? AND ?
    ORDER BY device_id, timestamp ASC
  `;

  db.query(query, [region_id, firstDay, lastDay], (err, results) => {
    if (err) {
      console.error("Error:", err);
      return res.status(500).json({ error: "Database error" });
    }

    // Filter out null values from each row
    const filteredResults = results.map(row => {
      const filteredRow = {};
      Object.keys(row).forEach(key => {
        if (row[key] !== null && row[key] !== undefined) {
          filteredRow[key] = row[key];
        }
      });
      return filteredRow;
    });

    res.json({
      status: "success",
      region_id,
      month: `${year}-${String(month).padStart(2, '0')}`,
      date_range: {
        from: firstDay,
        to: lastDay
      },
      total_records: filteredResults.length,
      data: filteredResults
    });
  });
};


// ==================== DEVICE LIST API ====================

const getDevices = (req, res) => {
  const { region_id } = req.query;

  let query = `
    SELECT DISTINCT device_id 
    FROM realtime_sensor_data 
    WHERE 1=1
  `;

  const params = [];

  if (region_id) {
    query += ` AND region_id = ?`;
    params.push(region_id);
  }

  query += ` ORDER BY device_id`;

  db.query(query, params, (err, results) => {
    if (err) {
      console.error("❌ Error fetching devices:", err);
      return res.status(500).json({ error: "Database error" });
    }

    const devices = results.map(row => row.device_id);
    console.log(`📱 Found ${devices.length} devices in region ${region_id || 'ALL'}`);
    res.json({ devices });
  });
};


const fetchDashboardDataby = (req, res) => {
  const { company, region } = req.query;

  if (!company || !region)
    return res.status(400).json({ error: 'Company and region are required' });

  const regionName = region.trim();

  // 1️⃣ Get devices in the region
  db.query(
    `
    SELECT d.device_id, d.region_id
    FROM devices d
    JOIN regions r ON d.region_id = r.region_id
    WHERE r.company_name = ? AND r.region_name = ?
    `,
    [company, regionName],
    (err, devices) => {
      if (err)
        return res.status(500).json({ error: 'DB error fetching devices' });

      if (!devices.length)
        return res.status(404).json({ error: 'No devices found' });

      const deviceIds = devices.map(d => d.device_id);
      const placeholders = deviceIds.map(() => '?').join(',');

      const results = [];
      let completed = 0;
      let hasRealtimeData = false;

      // 2️⃣ Fetch latest realtime data for each device using region_id
      devices.forEach((device) => {
        db.query(
          `
          SELECT *
          FROM realtime_sensor_data
          WHERE device_id = ? AND region_id = ?
          ORDER BY timestamp DESC
          LIMIT 1
          `,
          [device.device_id, device.region_id],
          (err, rows) => {
            completed++;

            if (!err && rows.length) {
              hasRealtimeData = true;

              const filteredRow = {};
              for (const key in rows[0]) {
                if (rows[0][key] !== null) {
                  filteredRow[key] = rows[0][key];
                }
              }

              results.push(filteredRow);
            }

            // 3️⃣ After all devices processed
            if (completed === devices.length) {

              // ✅ If realtime data exists → return it
              if (hasRealtimeData && results.length > 0) {
                return res.json({
                  status: 'success',
                  source: 'realtime_sensor_data',
                  company,
                  region: regionName,
                  devices: results
                });
              }

              // -------------------- FALLBACK TO DUMMY TABLE --------------------
              const dummyQuery = `
                SELECT *
                FROM dummy
                WHERE device_id IN (${placeholders})
                ORDER BY timestamp DESC
                LIMIT 6
              `;

              db.query(dummyQuery, deviceIds, (err, dummyResults) => {
                if (err)
                  return res.status(500).json({ error: 'DB error fetching dummy data' });

                return res.json({
                  status: 'success',
                  source: 'dummy',
                  company,
                  region: regionName,
                  devices: deviceIds,
                  data: dummyResults
                });
              });
            }
          }
        );
      });
    }
  );
};

// ==================== ANALYSIS ENDPOINT ====================

//final working code proper main code 
/*const generateAnalysisReport = (req, res) => {
  const {
    device_id,
    timeRange,
    shift,
    region_id
  } = req.query;

  if (!device_id || !timeRange) {
    return res.status(400).json({
      error: "device_id and timeRange required"
    });
  }

  console.log(`📊 Generating analysis for ${device_id} - ${timeRange} ${shift || ''} (Region: ${region_id || 'ALL'})`);

  // Check if we need ALL devices or a specific one
  const isAllDevices = device_id === 'all';

  // Choose the right data fetcher based on timeRange and device selection
  let dataFetcher;

  if (isAllDevices) {
    // Use the "ALL" versions of your functions
    if (timeRange === 'shift' && shift) {
      console.log('🔄 WRAPPER DEBUG:');
      console.log('  Original shift from query:', shift);
      console.log('  Original region_id:', region_id);

      let shiftParam = '';
      if (shift === '6am-2pm') shiftParam = 'morning';
      else if (shift === '2pm-10pm') shiftParam = 'afternoon';
      else if (shift === '10pm-6am') shiftParam = 'night';
      else shiftParam = shift;

      console.log('  Mapped shiftParam:', shiftParam);

      const newReq = {
        ...req,
        query: {
          ...req.query,
          shift: shiftParam,
          region_id: region_id
        }
      };

      console.log('  newReq.query:', newReq.query);

      dataFetcher = (req2, res2) => {
        console.log('📞 Calling getAllDevicesShiftData with:', req2.query);
        getAllDevicesShiftData(req2, res2);
      };
    }
    else if (timeRange === 'daily') {
      if (region_id) req.query.region_id = region_id;
      dataFetcher = getAllDevicesDailyData;
    }
    else if (timeRange === 'monthly') {
      if (region_id) req.query.region_id = region_id;
      dataFetcher = getAllDevicesMonthlyData;
    }
    else {
      return res.status(400).json({ error: "Invalid timeRange" });
    }
  } else {
    // Use single device versions
    if (timeRange === 'shift' && shift) {
      let shiftParam = '';
      if (shift === '6am-2pm') shiftParam = 'morning';
      else if (shift === '2pm-10pm') shiftParam = 'afternoon';
      else if (shift === '10pm-6am') shiftParam = 'night';

      req.query.shift = shiftParam;
      if (region_id) req.query.region_id = region_id;
      dataFetcher = getDeviceShiftData;
    }
    else if (timeRange === 'daily') {
      if (region_id) req.query.region_id = region_id;
      dataFetcher = getDeviceDailyData;
    }
    else if (timeRange === 'monthly') {
      if (region_id) req.query.region_id = region_id;
      dataFetcher = getDeviceMonthlyData;
    }
    else {
      return res.status(400).json({ error: "Invalid timeRange" });
    }
  }

  // Store the original res.json
  const originalJson = res.json;

  // Override res.json to capture the data
  res.json = function (data) {
    // Check if we have data in any format
    let records = [];

    // Handle different response formats
    if (data && data.data && Array.isArray(data.data)) {
      records = data.data;
    } else if (data && Array.isArray(data)) {
      records = data;
    } else if (data && data.results && Array.isArray(data.results)) {
      records = data.results;
    }

    console.log(`📊 Found ${records.length} records for analysis`);

    // Format data for Python
    const pythonInput = {
      data: records.map(row => ({
        device_id: row.device_id || device_id,
        time: row.timestamp,
        lat: parseFloat(row.latitude || 0),
        lon: parseFloat(row.longitude || 0),
        pitch: parseFloat(row.pitch || 0),
        roll: parseFloat(row.roll || 0),
        fuel: parseFloat(row.fuel || 0),
        fuel_consumption: parseFloat(row.fuel_consumption || 0),
        speed: parseFloat(row.speed || 0),
        distance: parseFloat(row.distance || 0),
        alt: parseFloat(row.altitude || 0),      // ✅ ADD altitude
        rl: parseFloat(row.rl || 0),
        fuel_cost: parseFloat(row.fuel_cost || 0),
        vibration: parseFloat(row.vibration || 0)
      })),
      report_type: timeRange
    };

    console.log(`🚀 Sending ${records.length} records to Python for analysis...`);

    // Call Python script
    //const pythonPath = '/opt/sample/venv/bin/python'; // ✅ venv python

    //const pythonProcess = exec(`${pythonPath} routes/analysis.py`, (error, stdout, stderr) => {
    const pythonProcess = exec('python routes/analysis.py', (error, stdout, stderr) => {
      if (error) {
        console.error('❌ Python error:', error);
        return originalJson.call(res, { error: "Analysis failed: " + error.message });
      }

      if (stderr) {
        console.log('📝 Python log:', stderr);
      }

      try {
        const result = JSON.parse(stdout);

        // Check the status from Python
        if (result.status === 'error') {
          console.error('❌ Python analysis error:', result.error);
          return originalJson.call(res, { error: result.error });
        }

        // Your analysis.py returns 'report' field with base64 Excel data
        if (!result.report) {
          console.error('❌ No report data in Python output');
          console.log('Python output keys:', Object.keys(result));
          return originalJson.call(res, { error: "No report data generated" });
        }

        // Decode the base64 Excel file
        const excelBuffer = Buffer.from(result.report, 'base64');

        // Use filename from Python
        const filename = result.filename || `analysis_${device_id}_${timeRange}_${new Date().toISOString().split('T')[0]}.xlsx`;

        // Set correct headers for Excel file download
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.setHeader('Content-Length', excelBuffer.length);
        res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
        res.setHeader('Pragma', 'no-cache');
        res.setHeader('Expires', '0');

        // Send the Excel file
        res.send(excelBuffer);

        console.log(`✅ Analysis complete! Excel report sent: ${filename}`);
        console.log(`📊 Report size: ${(excelBuffer.length / 1024).toFixed(2)} KB`);

      } catch (e) {
        console.error('❌ Failed to parse Python output:', e);
        console.log('Raw output (first 500 chars):', stdout.substring(0, 500));
        originalJson.call(res, { error: "Failed to generate report - invalid response from analysis engine" });
      }
    });

    pythonProcess.stdin.write(JSON.stringify(pythonInput));
    pythonProcess.stdin.end();
  };

  // Call the appropriate data fetcher
  dataFetcher(req, res);
};*/


//report for coustom data 
/*const generateAnalysisReport = (req, res) => {
  const {
    device_id,
    timeRange,
    shift,
    region_id
  } = req.query;

  if (!device_id || !timeRange) {
    return res.status(400).json({
      error: "device_id and timeRange required"
    });
  }

  // HARDCODED TO JUNE 30, 2026
  const targetDate = '2026-06-30';
  
  console.log(`📊 Generating analysis for ${device_id} - ${timeRange} (Date: ${targetDate}) ${shift || ''} (Region: ${region_id || 'ALL'})`);

  // Check if we need ALL devices or a specific one
  const isAllDevices = device_id === 'all';

  // Choose the right data fetcher based on timeRange and device selection
  let dataFetcher;

  if (isAllDevices) {
    // Use the "ALL" versions of your functions
    if (timeRange === 'shift' && shift) {
      console.log('🔄 WRAPPER DEBUG:');
      console.log('  Original shift from query:', shift);
      console.log('  Original region_id:', region_id);

      let shiftParam = '';
      if (shift === '6am-2pm') shiftParam = 'morning';
      else if (shift === '2pm-10pm') shiftParam = 'afternoon';
      else if (shift === '10pm-6am') shiftParam = 'night';
      else shiftParam = shift;

      console.log('  Mapped shiftParam:', shiftParam);

      // Directly modify req.query - HARDCODED DATE
      req.query.shift = shiftParam;
      req.query.region_id = region_id;

      console.log('  Modified req.query:', req.query);

      dataFetcher = getAllDevicesShiftData;
    }
    else if (timeRange === 'daily') {
      if (region_id) req.query.region_id = region_id;
      dataFetcher = getAllDevicesDailyData;
    }
    else if (timeRange === 'monthly') {
      if (region_id) req.query.region_id = region_id;
      dataFetcher = getAllDevicesMonthlyData;
    }
    else {
      return res.status(400).json({ error: "Invalid timeRange" });
    }
  } else {
    // Use single device versions
    if (timeRange === 'shift' && shift) {
      let shiftParam = '';
      if (shift === '6am-2pm') shiftParam = 'morning';
      else if (shift === '2pm-10pm') shiftParam = 'afternoon';
      else if (shift === '10pm-6am') shiftParam = 'night';

      req.query.shift = shiftParam;
      if (region_id) req.query.region_id = region_id;
      dataFetcher = getDeviceShiftData;
    }
    else if (timeRange === 'daily') {
      if (region_id) req.query.region_id = region_id;
      dataFetcher = getDeviceDailyData;
    }
    else if (timeRange === 'monthly') {
      if (region_id) req.query.region_id = region_id;
      dataFetcher = getDeviceMonthlyData;
    }
    else {
      return res.status(400).json({ error: "Invalid timeRange" });
    }
  }

  // Store the original res.json
  const originalJson = res.json;

  // Override res.json to capture the data
  res.json = function (data) {
    // Check if we have data in any format
    let records = [];

    // Handle different response formats
    if (data && data.data && Array.isArray(data.data)) {
      records = data.data;
    } else if (data && Array.isArray(data)) {
      records = data;
    } else if (data && data.results && Array.isArray(data.results)) {
      records = data.results;
    }

    console.log(`📊 Found ${records.length} records for analysis (Date: ${targetDate})`);

    // Format data for Python
    const pythonInput = {
      data: records.map(row => ({
        device_id: row.device_id || device_id,
        time: row.timestamp,
        lat: parseFloat(row.latitude || 0),
        lon: parseFloat(row.longitude || 0),
        pitch: parseFloat(row.pitch || 0),
        roll: parseFloat(row.roll || 0),
        fuel: parseFloat(row.fuel || 0),
        fuel_consumption: parseFloat(row.fuel_consumption || 0),
        speed: parseFloat(row.speed || 0),
        distance: parseFloat(row.distance || 0),
        alt: parseFloat(row.altitude || 0),
        rl: parseFloat(row.rl || 0),
        fuel_cost: parseFloat(row.fuel_cost || 0),
        vibration: parseFloat(row.vibration || 0)
      })),
      report_type: timeRange,
      target_date: targetDate
    };

    console.log(`🚀 Sending ${records.length} records to Python for analysis...`);

    // Call Python script
    const pythonProcess = exec('python routes/analysis.py', (error, stdout, stderr) => {
      if (error) {
        console.error('❌ Python error:', error);
        return originalJson.call(res, { error: "Analysis failed: " + error.message });
      }

      if (stderr) {
        console.log('📝 Python log:', stderr);
      }

      try {
        const result = JSON.parse(stdout);

        if (result.status === 'error') {
          console.error('❌ Python analysis error:', result.error);
          return originalJson.call(res, { error: result.error });
        }

        if (!result.report) {
          console.error('❌ No report data in Python output');
          console.log('Python output keys:', Object.keys(result));
          return originalJson.call(res, { error: "No report data generated" });
        }

        const excelBuffer = Buffer.from(result.report, 'base64');
        const filename = result.filename || `analysis_${device_id}_${timeRange}_${targetDate}.xlsx`;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.setHeader('Content-Length', excelBuffer.length);
        res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
        res.setHeader('Pragma', 'no-cache');
        res.setHeader('Expires', '0');

        res.send(excelBuffer);

        console.log(`✅ Analysis complete! Excel report sent: ${filename}`);
        console.log(`📊 Report size: ${(excelBuffer.length / 1024).toFixed(2)} KB`);

      } catch (e) {
        console.error('❌ Failed to parse Python output:', e);
        console.log('Raw output (first 500 chars):', stdout.substring(0, 500));
        originalJson.call(res, { error: "Failed to generate report - invalid response from analysis engine" });
      }
    });

    pythonProcess.stdin.write(JSON.stringify(pythonInput));
    pythonProcess.stdin.end();
  };

  // Call the appropriate data fetcher
  dataFetcher(req, res);
};
*/
const generateAnalysisReport = (req, res) => {
  const {
    timeRange,
    shift,
    company,
    sector,
    region
  } = req.query;

  // Validate required parameters
  if (!timeRange) {
    return res.status(400).json({
      error: "timeRange required"
    });
  }

  if (!company || !sector || !region) {
    return res.status(400).json({
      error: "company, sector, and region required"
    });
  }

  console.log(`📊 Generating analysis for Company: ${company}, Sector: ${sector}, Region: ${region}`);
  console.log(`📊 TimeRange: ${timeRange}${shift ? ', Shift: ' + shift : ''}`);

  // HARDCODED TO JUNE 30, 2026
  const targetDate = '2026-06-30';
  
  // Choose the right data fetcher based on timeRange
  let dataFetcher;

  if (timeRange === 'shift' && shift) {
    console.log('🔄 Shift report requested');
    console.log('  Shift from query:', shift);
    console.log('  Region:', region);
    console.log('  Company:', company);
    console.log('  Sector:', sector);

    let shiftParam = '';
    if (shift === '6am-2pm') shiftParam = 'morning';
    else if (shift === '2pm-10pm') shiftParam = 'afternoon';
    else if (shift === '10pm-6am') shiftParam = 'night';
    else shiftParam = shift;

    console.log('  Mapped shiftParam:', shiftParam);

    // Pass all parameters to the data fetcher
    req.query.shift = shiftParam;
    req.query.company = company;
    req.query.sector = sector;
    req.query.region = region;

    dataFetcher = getAllDevicesShiftData;
  }
  else if (timeRange === 'daily') {
    req.query.company = company;
    req.query.sector = sector;
    req.query.region = region;
    dataFetcher = getAllDevicesDailyData;
  }
  else if (timeRange === 'monthly') {
    req.query.company = company;
    req.query.sector = sector;
    req.query.region = region;
    dataFetcher = getAllDevicesMonthlyData;
  }
  else {
    return res.status(400).json({ error: "Invalid timeRange" });
  }

  // Store the original res.json
  const originalJson = res.json;

  // Override res.json to capture the data
  res.json = function (data) {
    // Check if we have data in any format
    let records = [];

    // Handle different response formats
    if (data && data.data && Array.isArray(data.data)) {
      records = data.data;
    } else if (data && Array.isArray(data)) {
      records = data;
    } else if (data && data.results && Array.isArray(data.results)) {
      records = data.results;
    }

    console.log(`📊 Found ${records.length} records for analysis (Date: ${targetDate})`);

    if (records.length === 0) {
      console.log('⚠️ No records found, returning empty report');
      return originalJson.call(res, { 
        status: 'warning', 
        message: 'No data found for the selected criteria'
      });
    }

    // Format data for Python
    const pythonInput = {
      data: records.map(row => ({
        device_id: row.device_id || 'unknown',
        time: row.timestamp,
        lat: parseFloat(row.latitude || 0),
        lon: parseFloat(row.longitude || 0),
        pitch: parseFloat(row.pitch || 0),
        roll: parseFloat(row.roll || 0),
        fuel: parseFloat(row.fuel || 0),
        fuel_consumption: parseFloat(row.fuel_consumption || 0),
        speed: parseFloat(row.speed || 0),
        distance: parseFloat(row.distance || 0),
        alt: parseFloat(row.altitude || 0),
        rl: parseFloat(row.rl || 0),
        fuel_cost: parseFloat(row.fuel_cost || 0),
        vibration: parseFloat(row.vibration || 0)
      })),
      report_type: timeRange,
      target_date: targetDate,
      shift: shift || null,
      company: company,
      sector: sector,
      region: region
    };

    console.log(`🚀 Sending ${records.length} records to Python for analysis...`);

    // Call Python script
    const pythonProcess = exec('python routes/analysis.py', (error, stdout, stderr) => {
      if (error) {
        console.error('❌ Python error:', error);
        return originalJson.call(res, { error: "Analysis failed: " + error.message });
      }

      if (stderr) {
        console.log('📝 Python log:', stderr);
      }

      try {
        const result = JSON.parse(stdout);

        if (result.status === 'error') {
          console.error('❌ Python analysis error:', result.error);
          return originalJson.call(res, { error: result.error });
        }

        if (!result.report) {
          console.error('❌ No report data in Python output');
          console.log('Python output keys:', Object.keys(result));
          return originalJson.call(res, { error: "No report data generated" });
        }

        const excelBuffer = Buffer.from(result.report, 'base64');
        const filename = result.filename || `analysis_${company}_${region}_${timeRange}_${targetDate}.xlsx`;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.setHeader('Content-Length', excelBuffer.length);
        res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
        res.setHeader('Pragma', 'no-cache');
        res.setHeader('Expires', '0');

        res.send(excelBuffer);

        console.log(`✅ Analysis complete! Excel report sent: ${filename}`);
        console.log(`📊 Report size: ${(excelBuffer.length / 1024).toFixed(2)} KB`);

      } catch (e) {
        console.error('❌ Failed to parse Python output:', e);
        console.log('Raw output (first 500 chars):', stdout.substring(0, 500));
        originalJson.call(res, { error: "Failed to generate report - invalid response from analysis engine" });
      }
    });

    pythonProcess.stdin.write(JSON.stringify(pythonInput));
    pythonProcess.stdin.end();
  };

  // Call the appropriate data fetcher
  dataFetcher(req, res);
};
// ==================== EXPORT ALL FUNCTIONS ====================
module.exports = {
  register,
  signin,
  forgotPassword,

  insertRealtimeData,  // ✅ UPDATED: Calculates and stores ALL fields
  getLast10ZAxis,     //fuel and gradient analysis chart in kacha 
  registerToken,

  getDeviceShiftData,
  getAllDevicesShiftData,
  getDeviceDailyData,
  getAllDevicesDailyData,
  getDeviceMonthlyData,
  getAllDevicesMonthlyData,

  generateAnalysisReport,  // ✅ ADD THIS
  getDevices,
  fetchDashboardDataby
};