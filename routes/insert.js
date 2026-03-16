const ExcelJS = require('exceljs');
const pool = require('../dao/dao');
const db = require('../dao/dao');
const nodemailer = require('nodemailer');
const { exec } = require('child_process');

// ==================== GLOBAL CONSTANTS ====================
const DIESEL_PRICE_PER_LITER = 94.5; // ₹ per liter
const SEA_LEVEL_RL = 525.5; // Fixed sea level height

// ==================== UTILITY FUNCTIONS ====================

// Haversine formula to calculate distance between two coordinates in kilometers
function haversineKm(coord1, coord2) {
  const toRad = x => (x * Math.PI) / 180;
  const R = 6371; // Earth's radius in km
  const dLat = toRad(coord2[0] - coord1[0]);
  const dLon = toRad(coord2[1] - coord1[1]);
  const lat1 = toRad(coord1[0]);
  const lat2 = toRad(coord2[0]);
  const a =
    Math.sin(dLat / 2) ** 2 +
    Math.sin(dLon / 2) ** 2 * Math.cos(lat1) * Math.cos(lat2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c; // Distance in km
}

// ==================== STATUS CALCULATION FUNCTIONS ====================

// Function to get movement label from pitch
function getMovementFromPitch(pitch) {
  if (pitch === null || pitch === undefined || isNaN(pitch)) {
    return 'STABLE';
  }
  
  const pitchValue = Number(pitch);
  if (pitchValue > 5) return 'UPHILL';
  if (pitchValue < -5) return 'DOWNHILL';
  return 'STABLE';
}

// Function to get vibration label
function getVibrationLabel(vibrationValue) {
  const vib = Number(vibrationValue);
  if (Number.isNaN(vib)) return 'UNKNOWN';
  
  switch(vib) {
    case 1: return 'VERY LOW';
    case 2: return 'LOW';
    case 3: return 'MIDDLE';
    case 4: return 'HIGH';
    case 5: return 'VERY HIGH';
    default: return 'UNKNOWN';
  }
}

// Function to detect excavator state
function detectExcavatorState(vibrationValue, movementValue) {
  const vib = Number(vibrationValue);
  const movement = movementValue || 'STABLE';
  
  if (Number.isNaN(vib)) return 'UNKNOWN';
  
  if (vib <= 2) {
    return 'IDLE';
  } else if (vib === 3) {
    return 'OPERATING';
  } else if (vib >= 4) {
    return 'LOADING';
  }
  
  return 'UNKNOWN';
}

// Function to get segment type from pitch
function getSegmentType(pitch) {
  const pitchValue = Math.abs(Number(pitch) || 0);
  return pitchValue <= 5 ? 'FLAT' : 'GRADIENT';
}

// Gradient multiplier function (based on pitch)
function getGradientMultiplier(pitch = 0) {
  const pitchValue = Number(pitch);
  if (pitchValue <= -5) return 0.25;        // Downhill
  if (pitchValue > -5 && pitchValue <= 3) return 0.65; // Flat
  if (pitchValue > 3 && pitchValue <= 8) return 1.3;   // Mild uphill
  return 2.0;                          // Steep uphill
}

// Speed multiplier function
function getSpeedMultiplier(speed = 0) {
  const speedValue = Number(speed);
  if (speedValue <= 5) return 0.9;      // Idle / slow
  if (speedValue <= 20) return 1.0;     // Normal
  if (speedValue <= 35) return 1.1;     // Loaded
  return 1.25;                     // Overspeed / stress
}

// Calculate RL (Reduced Level)
function calculateRL(altitude) {
  if (altitude === null || altitude === undefined || isNaN(altitude)) {
    return null;
  }
  return Number((Number(altitude) + SEA_LEVEL_RL).toFixed(2));
}

// ==================== FUEL AND COST CALCULATION ====================

// FIXED calculateFuelAndCost function for D7:
function calculateFuelAndCost(distance, pitch, movement, deviceId, timeDiffHours = 0) {
  let fuel = 0;
  
  // For very small distances (less than 1 meter), use minimal fuel
  if (distance < 0.001) { // < 1 meter
    return {
      fuel: 0.000001, // 0.001 mL
      cost: 0.00
    };
  }
  
  // Check if device actually moved (more than 1 meter)
  const isMoving = distance > 0.001; // > 1 meter
  
  // ALL DEVICES use distance-based when MOVING
  if (isMoving && distance > 0) {
    // Base: 1.52 km per liter = 0.6579 liters per km
    let litersPerKm = 1 / 1.52; // ~0.6579 L/km
    
    // Base fuel for this segment
    fuel = distance * litersPerKm; // distance in km
    
    // Apply gradient multiplier based on pitch
    const gradientMultiplier = getGradientMultiplier(pitch);
    
    // movement is already numeric (converted from string)
    const speedMultiplier = getSpeedMultiplier(movement);
    
    fuel = fuel * gradientMultiplier * speedMultiplier;
    
    //console.log(`🚚 Distance-based fuel for ${(distance*1000).toFixed(1)}m: ${(fuel*1000).toFixed(1)}mL`);
  }
  // Only use time-based for STATIONARY excavator
  else if (deviceId === 'D7' && !isMoving) {
    // Cap timeDiffHours to reasonable maximum
    const MAX_TIME_GAP_HOURS = 1.0; // Maximum 1 hour between readings
    const realisticTimeDiff = Math.min(timeDiffHours, MAX_TIME_GAP_HOURS);
    
    // 15 liters per hour (only when NOT moving)
    fuel = 15 * realisticTimeDiff;
    
    console.log(`🏗️ Stationary excavator time: ${timeDiffHours.toFixed(2)}h → ${realisticTimeDiff.toFixed(2)}h = ${fuel.toFixed(3)}L`);
  }
  // Minimal fuel for stationary non-excavator
  else {
    fuel = 0.00001; // 0.01 mL minimal fuel
  }
  
  const cost = fuel * DIESEL_PRICE_PER_LITER;
  
  return {
    fuel: parseFloat(fuel.toFixed(6)),
    cost: parseFloat(cost.toFixed(2))
  };
}

//excavator no need fuel and fuelcodt
/*function calculateFuelAndCost(distance, pitch, movement, equipment_name) {

  let fuel = 0;

  // 1️⃣ Excavators → no fuel calculation
  if (equipment_name.toLowerCase() === "excavator") {
    return {
      fuel: 0,
      cost: 0
    };
  }

  // 2️⃣ Ignore GPS drift (<1 meter)
  if (distance < 0.001) {
    return {
      fuel: 0,
      cost: 0
    };
  }

  const isMoving = distance > 0.001;

  // 3️⃣ Hauler fuel calculation
  if (equipment_name.toLowerCase() === "hauler" && isMoving) {

    // Base fuel rate (1.52 km per liter)
    const litersPerKm = 1 / 1.52;

    fuel = distance * litersPerKm;

    const gradientMultiplier = getGradientMultiplier(pitch);
    const movementMultiplier = getMovementMultiplier(movement);

    fuel = fuel * gradientMultiplier * movementMultiplier;

  }

  const cost = fuel * DIESEL_PRICE_PER_LITER;

  return {
    fuel: Number(fuel.toFixed(6)),
    cost: Number(cost.toFixed(2))
  };
}*/

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

//hari 2. INSERT REAL-TIME DATA WITH ALL CALCULATIONS AND REGION_ID
/*const insertRealtimeData = (req, res) => {
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
    vibration
    // REMOVED timestamp - we don't accept client timestamp
  } = req.body;

  if (!device_id) {
    return res.status(400).json({ error: "Missing required field: device_id" });
  }

  // Get region_id from devices table
  const getRegionQuery = `SELECT region_id FROM devices WHERE device_id = ?`;

  db.query(getRegionQuery, [device_id], (err, regionResults) => {
    if (err) {
      console.error("❌ Error fetching region_id:", err);
      return res.status(500).json({ error: "Database error fetching region_id" });
    }

    if (regionResults.length === 0) {
      console.error(`❌ Device ${device_id} not found in devices table`);
      return res.status(404).json({ error: `Device ${device_id} not found` });
    }

    const region_id = regionResults[0].region_id;
    console.log(`📍 Device ${device_id} belongs to region_id: ${region_id}`);

    // Get previous point for distance calculation
    const getPreviousPointQuery = `
      SELECT latitude, longitude, timestamp 
      FROM realtime_sensor_data 
      WHERE device_id = ? 
      ORDER BY id DESC 
      LIMIT 1
    `;

    db.query(getPreviousPointQuery, [device_id], (err, prevResults) => {
      if (err) {
        console.error("❌ Error fetching previous point:", err);
        return res.status(500).json({ error: "Database error fetching previous data" });
      }

      let distance = 0;
      let timeDiffHours = 0;

      // Calculate distance from previous point
      if (prevResults.length > 0 && prevResults[0].latitude && prevResults[0].longitude) {
        try {
          const prevLat = parseFloat(prevResults[0].latitude);
          const prevLon = parseFloat(prevResults[0].longitude);
          const currLat = parseFloat(latitude);
          const currLon = parseFloat(longitude);
          
          distance = haversineKm([prevLat, prevLon], [currLat, currLon]);
          
          // Calculate time difference using MySQL timestamps
          const prevTime = new Date(prevResults[0].timestamp);
          const currentTime = new Date();
          timeDiffHours = Math.max(0, (currentTime - prevTime) / 3600000);
          
          console.log(`📏 Segment distance: ${(distance * 1000).toFixed(2)} m`);
          console.log(`⏱️ Time since last: ${(timeDiffHours * 3600).toFixed(1)} sec`);
        } catch (error) {
          console.error('❌ Error in calculation:', error);
        }
      } else {
        console.log('📌 First data point - distance = 0');
      }

      // Convert movement string to numeric
      let movementNumeric = 0;
      if (movement) {
        if (movement === 'DOWN' || movement === 'DOWNHILL') movementNumeric = -10;
        else if (movement === 'UP' || movement === 'UPHILL') movementNumeric = 10;
        else if (movement === 'STABLE' || movement === 'FLAT') movementNumeric = 0;
        else movementNumeric = parseFloat(movement) || 0;
      }

      // Calculate fuel and cost
      const segmentFuelResult = calculateFuelAndCost(
        distance,
        parseFloat(pitch) || 0,
        movementNumeric,
        device_id,
        timeDiffHours
      );

      // Calculate RL
      const rl = altitude !== undefined ? (parseFloat(altitude) + SEA_LEVEL_RL).toFixed(2) : null;

      // ===== CORRECT INSERT with IST timestamp =====
      // Set MySQL session to IST before inserting
      const setTimezoneQuery = "SET SESSION time_zone = '+05:30'";
      
      db.query(setTimezoneQuery, (timezoneErr) => {
        if (timezoneErr) {
          console.warn("⚠️ Could not set timezone, using default:", timezoneErr);
        }

        // INSERT query - MySQL will use IST now
        const insertQuery = `
          INSERT INTO realtime_sensor_data (
            device_id,
            equipment_name,
            timestamp,  -- MySQL will use current IST time
            latitude,
            longitude,
            altitude,
            speed,
            pitch,
            roll,
            movement,
            vibration,
            distance,
            fuel,
            fuel_cost,
            rl,
            region_id 
          )
          VALUES (?, ?, NOW(), ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `;

        const values = [
          device_id,
          latitude ? parseFloat(latitude) : null,
          longitude ? parseFloat(longitude) : null,
          altitude !== undefined ? parseFloat(altitude) : null,
          z_axis !== undefined ? parseFloat(z_axis) : null,
          movement || null,
          pitch !== undefined ? parseInt(pitch) : null,
          roll !== undefined ? parseInt(roll) : null,
          vibration !== undefined ? parseInt(vibration) : null,
          parseFloat(distance.toFixed(6)),
          parseFloat(segmentFuelResult.fuel.toFixed(6)),
          parseFloat(segmentFuelResult.cost.toFixed(2)),
          rl,
          region_id
        ];

        console.log('📝 Storing calculated values:');
        console.log(`   Device ID: ${device_id}`);
        console.log(`   Region ID: ${region_id}`);
        console.log(`   Distance: ${(distance * 1000).toFixed(2)} m`);
        console.log(`   Fuel: ${(segmentFuelResult.fuel * 1000).toFixed(2)} mL`);
        console.log(`   Cost: ₹${segmentFuelResult.cost.toFixed(4)}`);
        console.log(`   RL: ${rl || 'N/A'} m`);
        console.log(`   Movement: ${movement || 'N/A'}`);

        db.query(insertQuery, values, (err, result) => {
          if (err) {
            console.error("❌ Database insert error:", err.sqlMessage);
            return res.status(500).json({ error: "Database error: " + err.message });
          }

          // Verify the IST time that was stored
          const verifyQuery = "SELECT NOW() as ist_time";
          
          db.query(verifyQuery, (timeErr, timeResult) => {
            const istTime = timeErr ? 'unknown' : timeResult[0].ist_time;
            
            console.log(`\n✅ Stored data for ${device_id}:`);
            console.log(`   ID: ${result.insertId}`);
            console.log(`   Region ID: ${region_id}`);
            console.log(`   IST Time: ${istTime}`);
            console.log(`   Position: ${parseFloat(latitude).toFixed(6)}, ${parseFloat(longitude).toFixed(6)}`);
            console.log(`   Altitude: ${altitude || 'N/A'} m`);
            console.log(`   RL: ${rl || 'N/A'} m`);
            console.log(`   Segment Distance: ${(distance * 1000).toFixed(2)} m`);
            console.log(`   Segment Fuel: ${(segmentFuelResult.fuel * 1000).toFixed(2)} mL`);
            console.log(`   Segment Cost: ₹${segmentFuelResult.cost.toFixed(4)}`);
            console.log('─────────────────────────────────────────────');

            res.json({
              status: "success",
              message: "Data stored with IST timestamp",
              inserted_id: result.insertId,
              region_id: region_id,
              ist_timestamp: istTime,
              calculations: {
                distance_km: parseFloat(distance.toFixed(6)),
                distance_m: parseFloat((distance * 1000).toFixed(2)),
                fuel_l: parseFloat(segmentFuelResult.fuel.toFixed(6)),
                fuel_ml: parseFloat((segmentFuelResult.fuel * 1000).toFixed(2)),
                cost: parseFloat(segmentFuelResult.cost.toFixed(4)),
                rl: rl
              }
            });
          });
        });
      });
    });
  });
};*/

/*//THIS CODE WAS WORKING GOOD 
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
    vibration
  } = req.body;

  if (!device_id) {
    return res.status(400).json({ error: "Missing required field: device_id" });
  }

  // Get region_id from devices table
  const getRegionQuery = `SELECT region_id FROM devices WHERE device_id = ?`;

  db.query(getRegionQuery, [device_id], (err, regionResults) => {
    if (err) {
      console.error("❌ Error fetching region_id:", err);
      return res.status(500).json({ error: "Database error fetching region_id" });
    }

    if (regionResults.length === 0) {
      console.error(`❌ Device ${device_id} not found in devices table`);
      return res.status(404).json({ error: `Device ${device_id} not found` });
    }

    const region_id = regionResults[0].region_id;

    // Get previous point for distance calculation
    const getPreviousPointQuery = `
      SELECT latitude, longitude, timestamp 
      FROM realtime_sensor_data 
      WHERE device_id = ? 
      ORDER BY id DESC 
      LIMIT 1
    `;

    db.query(getPreviousPointQuery, [device_id], (err, prevResults) => {
      let distance = 0;
      let timeDiffHours = 0;

      if (err) {
        console.error("❌ Error fetching previous point:", err);
        return res.status(500).json({ error: "Database error fetching previous data" });
      }
      else {
      // Calculate distance from previous point
      if (prevResults.length > 0 && prevResults[0].latitude && prevResults[0].longitude) {
        try {
          const prevLat = parseFloat(prevResults[0].latitude);
          const prevLon = parseFloat(prevResults[0].longitude);
          const currLat = parseFloat(latitude);
          const currLon = parseFloat(longitude);
          
          distance = haversineKm([prevLat, prevLon], [currLat, currLon]);
          
          const prevTime = new Date(prevResults[0].timestamp);
          const currentTime = new Date();
          timeDiffHours = Math.max(0, (currentTime - prevTime) / 3600000);
        } catch (error) {
          console.error('❌ Error in calculation:', error);
        }
      }
      }
      // Convert movement string to numeric
      let movementNumeric = 0;
      if (movement) {
        if (movement === 'DOWN' || movement === 'DOWNHILL') movementNumeric = -10;
        else if (movement === 'UP' || movement === 'UPHILL') movementNumeric = 10;
        else if (movement === 'STABLE' || movement === 'FLAT') movementNumeric = 0;
        else movementNumeric = parseFloat(movement) || 0;
      }

      // Calculate fuel and cost
      const segmentFuelResult = calculateFuelAndCost(
        distance,
        parseFloat(pitch) || 0,
        movementNumeric,
        device_id,
        timeDiffHours
      );

      // Calculate RL
      const rl = altitude !== undefined ? (parseFloat(altitude) + SEA_LEVEL_RL).toFixed(2) : null;

      // Set MySQL session to IST
      const setTimezoneQuery = "SET SESSION time_zone = '+05:30'";
      
      db.query(setTimezoneQuery, (timezoneErr) => {
        if (timezoneErr) {
          console.warn("⚠️ Could not set timezone:", timezoneErr);
        }

        // INSERT query
        const insertQuery = `
          INSERT INTO realtime_sensor_data (
            device_id, equipment_name, timestamp, latitude, longitude, altitude,
            speed, pitch, roll, movement, vibration, distance, fuel, fuel_cost, rl, region_id
          ) VALUES (?, ?, NOW(), ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `;

        const values = [
          device_id,
          equipment_name,
          latitude ? parseFloat(latitude) : null,
          longitude ? parseFloat(longitude) : null,
          altitude !== undefined ? parseFloat(altitude) : null,
          speed !== undefined ? parseFloat(speed) : null,
          pitch !== undefined ? parseFloat(pitch) : null,
          roll !== undefined ? parseFloat(roll) : null,
          movement || null,
          vibration !== undefined ? parseFloat(vibration) : null,
          parseFloat(distance.toFixed(6)),
          parseFloat(segmentFuelResult.fuel.toFixed(6)),
          parseFloat(segmentFuelResult.cost.toFixed(4)),
          rl,
          region_id
        ];

        db.query(insertQuery, values, (err, result) => {
          if (err) {
            console.error("❌ Database insert error:", err.sqlMessage);
            return res.status(500).json({ error: "Database error: " + err.message });
          }

          // FINAL RESULT PRINT - Simple and clean
          console.log('\n✅ FINAL RESULT:');
          console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
          console.log(`Device ID: ${device_id}`);
          console.log(`Equipment: ${equipment_name || 'N/A'}`);
          console.log(`Position: ${latitude}, ${longitude}`);
          console.log(`Altitude: ${altitude || 0} m`);
          console.log(`Speed: ${speed || 0}`);
          console.log(`Pitch: ${pitch || 0}`);
          console.log(`Roll: ${roll || 0}`);
          console.log(`Movement: ${movement || 'N/A'}`);
          console.log(`Vibration: ${vibration || 0}`);
          console.log(`RL: ${rl || 0} m`);
          console.log(`Distance: ${(distance * 1000).toFixed(2)} m`);
          console.log(`Fuel: ${(segmentFuelResult.fuel * 1000).toFixed(2)} mL`);
          console.log(`Fuel Cost: ₹${segmentFuelResult.cost.toFixed(4)}`);
          console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');

          res.json({
            status: "success",
            message: "Data stored successfully",
            inserted_id: result.insertId
          });
        });
      });
    });
  });
};*/

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
    fuel,        // NEW
    pressure,
    count1     // NEW
  } = req.body;

  if (!device_id) {
    return res.status(400).json({ error: "Missing required field: device_id" });
  }

  const FUEL_PRICE_PER_LITER = 90;

  // Get region_id from devices table
  const getRegionQuery = `SELECT region_id FROM devices WHERE device_id = ?`;

  db.query(getRegionQuery, [device_id], (err, regionResults) => {

    if (err) {
      console.error("❌ Error fetching region_id:", err);
      return res.status(500).json({ error: "Database error fetching region_id" });
    }

    if (regionResults.length === 0) {
      console.error(`❌ Device ${device_id} not found`);
      return res.status(404).json({ error: `Device ${device_id} not found` });
    }

    const region_id = regionResults[0].region_id;

    // Get previous GPS point
    const getPreviousPointQuery = `
      SELECT latitude, longitude, timestamp
      FROM realtime_sensor_data
      WHERE device_id = ?
      ORDER BY id DESC
      LIMIT 1
    `;

    db.query(getPreviousPointQuery, [device_id], (err, prevResults) => {

      let distance = 0;
      let timeDiffHours = 0;

      if (err) {
        console.error("❌ Error fetching previous data:", err);
        return res.status(500).json({ error: "Database error fetching previous data" });
      }

      // Distance calculation
      if (prevResults.length > 0 && prevResults[0].latitude && prevResults[0].longitude) {
        try {

          const prevLat = parseFloat(prevResults[0].latitude);
          const prevLon = parseFloat(prevResults[0].longitude);

          const currLat = parseFloat(latitude);
          const currLon = parseFloat(longitude);

          distance = haversineKm([prevLat, prevLon], [currLat, currLon]);

          const prevTime = new Date(prevResults[0].timestamp);
          const currentTime = new Date();

          timeDiffHours = Math.max(0, (currentTime - prevTime) / 3600000);

        } catch (error) {
          console.error("❌ Error in distance calculation:", error);
        }
      }

      // Movement numeric conversion
      let movementNumeric = 0;

      if (movement) {
        if (movement === "DOWN" || movement === "DOWNHILL") movementNumeric = -10;
        else if (movement === "UP" || movement === "UPHILL") movementNumeric = 10;
        else if (movement === "STABLE" || movement === "FLAT") movementNumeric = 0;
        else movementNumeric = parseFloat(movement) || 0;
      }

      // -------------------------
      // Fuel Logic
      // -------------------------

      let fuelUsed = 0;
      let fuelCost = 0;

      if (fuel !== undefined && fuel !== null) {

        // Fuel sensor value
        fuelUsed = parseFloat(fuel);

        console.log("⛽ Using Fuel Sensor Value");

      } else {

        // Calculate fuel
        const calcResult = calculateFuelAndCost(
          distance,
          parseFloat(pitch) || 0,
          movementNumeric,
          device_id,
          timeDiffHours
        );

        fuelUsed = calcResult.fuel;

        console.log("⚙️ Using Calculated Fuel");
      }

      // Cost for both cases
      fuelCost = fuelUsed * FUEL_PRICE_PER_LITER;

      // RL calculation
      const rl =
        altitude !== undefined
          ? (parseFloat(altitude) + SEA_LEVEL_RL).toFixed(2)
          : null;

      // Set MySQL timezone
      const setTimezoneQuery = "SET SESSION time_zone = '+05:30'";

      db.query(setTimezoneQuery, (timezoneErr) => {

        if (timezoneErr) {
          console.warn("⚠️ Could not set timezone:", timezoneErr);
        }

        // Insert query
        const insertQuery = `
          INSERT INTO realtime_sensor_data (
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
            count1
          ) VALUES (?, ?, NOW(), ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `;

        const values = [
          device_id,
          equipment_name || null,
          latitude ? parseFloat(latitude) : null,
          longitude ? parseFloat(longitude) : null,
          altitude !== undefined ? parseFloat(altitude) : null,
          speed !== undefined ? parseFloat(speed) : null,
          pitch !== undefined ? parseFloat(pitch) : null,
          roll !== undefined ? parseFloat(roll) : null,
          movement || null,
          vibration !== undefined ? parseFloat(vibration) : null,
          pressure !== undefined ? parseFloat(pressure) : null,
          parseFloat(distance.toFixed(6)),
          parseFloat(fuelUsed.toFixed(6)),
          parseFloat(fuelCost.toFixed(4)),
          rl,
          region_id,
          count1 || null
        ];

        db.query(insertQuery, values, (err, result) => {

          if (err) {
            console.error("❌ Database insert error:", err.sqlMessage);
            return res.status(500).json({ error: "Database error: " + err.message });
          }

          // Console output
          console.log("\n✅ FINAL RESULT:");
          console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");

          console.log(`Device ID: ${device_id}`);
          console.log(`Equipment: ${equipment_name || "N/A"}`);

          console.log(`Position: ${latitude}, ${longitude}`);
          console.log(`Altitude: ${altitude || 0} m`);

          console.log(`Speed: ${speed || 0}`);
          console.log(`Pitch: ${pitch || 0}`);
          console.log(`Roll: ${roll || 0}`);

          console.log(`Movement: ${movement || "N/A"}`);
          console.log(`Vibration: ${vibration || 0}`);
          console.log(`Pressure: ${pressure || 0}`);

          console.log(`RL: ${rl || 0} m`);

          console.log(`Distance: ${(distance * 1000).toFixed(2)} m`);

          console.log(`Fuel Used: ${(fuelUsed * 1000).toFixed(2)} mL`);
          console.log(`Fuel Cost: ₹${fuelCost.toFixed(4)}`);
           console.log(`dump count: ${count1}`);

          console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");

          res.json({
            status: "success",
            message: "Data stored successfully",
            inserted_id: result.insertId
          });

        });

      });

    });

  });

};
// 3. FETCH DASHBOARD DATA (Updated to include calculated fields)
const fetchDashboardData = (req, res) => {
  const { company, region } = req.query;
  
  if (!company || !region) {
    return res.status(400).json({ error: 'Company and region are required' });
  }

  const regionName = region.trim();
  const isKache = regionName.toLowerCase() === 'kache';

  if (isKache) {
    let regionIds = [];
    const parsed = parseInt(region);

    if (!isNaN(parsed)) {
      regionIds = [parsed];
      proceedWithKacheQuery();
    } else {
      const regionQuery = `
        SELECT region_id FROM regions 
        WHERE region_name = ? AND company_name = ?
      `;
      db.query(regionQuery, [regionName, company], (err, results) => {
        if (err) return res.status(500).json({ error: 'DB error fetching region ID' });
        if (!results.length) return res.status(404).json({ error: 'Region not found' });
        regionIds = results.map(r => r.region_id);
        proceedWithKacheQuery();
      });
    }

    function proceedWithKacheQuery() {
      const deviceQuery = `
        SELECT DISTINCT UPPER(d.device_id) AS device_id
        FROM realtime_sensor_data d
        JOIN devices dev ON UPPER(d.device_id) = UPPER(dev.device_id)
        JOIN regions r ON dev.region_id = r.region_id
        WHERE r.company_name = ?
          AND r.region_id IN (?)
          AND d.device_id IN ('D3','D7','D8','D9','D12')
      `;

      db.query(deviceQuery, [company, regionIds], (err, devices) => {
        if (err) return res.status(500).json({ error: 'DB error fetching devices' });
        const deviceIds = devices.map(d => d.device_id);
        if (!deviceIds.length) return res.status(404).json({ error: 'No devices found' });

        const placeholders = deviceIds.map(() => '?').join(',');

        // Get latest reading for each device - ONLY FROM realtime_sensor_data
        const latestQuery = `
          SELECT 
            device_id,
            timestamp,
            latitude,
            longitude,
            altitude,
            z_axis,
            movement,
            pitch,
            roll,
            vibration,
            distance,
            fuel,
            fuel_cost,
            rl
          FROM (
            SELECT *,
                   ROW_NUMBER() OVER (PARTITION BY device_id ORDER BY id DESC) rn
            FROM realtime_sensor_data
            WHERE device_id IN (${placeholders})
          ) t
          WHERE rn = 1
        `;

        db.query(latestQuery, deviceIds, (err, latestRows) => {
          if (err) return res.status(500).json({ error: 'DB error fetching latest data' });

          const sites = {};
          const haulers = [];

          latestRows.forEach(row => {
            // Prepare site data
            sites[row.device_id] = {
              id: row.device_id,
              timestamp: row.timestamp,
              pos: [row.latitude, row.longitude],
              rl: row.rl,
              altitude: row.altitude,
              pitch: row.pitch,
              roll: row.roll,
              movement: row.movement,
              vibration: row.vibration,
              z_axis: row.z_axis,
              distance: row.distance,
              fuel: row.fuel,
              fuel_cost: row.fuel_cost
            };

            // Prepare hauler data
           haulers.push({
              id: row.device_id,
              timestamp: row.timestamp,
              pitch: row.pitch,
              roll: row.roll,
              altitude: row.altitude,
              movement: row.movement,
              vibration: row.vibration,
              distance: row.distance,
              fuel: row.fuel,
              fuel_cost: row.fuel_cost,
              rl: row.rl,
              z_axis: row.z_axis
            });
          });

          return res.json({
            status: 'success',
            company,
            region: regionIds,
            sites,
            //haulers
          });
        });
      });
    }
    return;
  }

  // For other regions
  const deviceQuery = `
    SELECT d.device_id
    FROM devices d
    JOIN regions r ON d.region_id = r.region_id
    WHERE r.company_name = ?
      AND r.region_name = ?
  `;

  db.query(deviceQuery, [company, regionName], (err, devices) => {
    if (err) return res.status(500).json({ error: 'DB error fetching devices' });
    if (!devices.length) return res.status(404).json({ error: 'No devices found' });

    const deviceIds = devices.map(d => d.device_id);
    const placeholders = deviceIds.map(() => '?').join(',');

    const sensorQuery = `
      SELECT *
      FROM dummy
      WHERE device_id IN (${placeholders})
      ORDER BY timestamp DESC
      LIMIT 6
    `;

    db.query(sensorQuery, deviceIds, (err, sensorResults) => {
      if (err) return res.status(500).json({ error: 'DB error fetching dummy data' });

      return res.json({
        status: 'success',
        company,
        region: regionName,
        devices: deviceIds,
        data: sensorResults
      });
    });
  });
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

// Receive sensor data endpoint (unchanged)
/*const receiveSensorData = (req, res) => {
  const { device_id, temperature, humidity, dust } = req.body;

  if (
    device_id === undefined ||
    temperature === undefined ||
    humidity === undefined ||
    dust === undefined
  ) {
    return res.status(400).json({ error: 'Missing sensor data' });
  }

  const insertQuery = `
    INSERT INTO dummy (device_id, temperature, humidity, dust, timestamp)
    VALUES (?, ?, ?, ?, NOW())
  `;

  db.query(insertQuery, [device_id, temperature, humidity, dust], (err, result) => {
    if (err) return res.status(500).json({ error: 'SQL insert failed' });

    return res.status(201).json({
      message: 'Sensor data stored successfully',
      id: result.insertId,
    });
  });
};*/

// Get last 10 z-axis values endpoint (unchanged)
const getLast10ZAxis = (req, res) => {
  const query = `
    SELECT device_id, pitch AS z_axis, timestamp
    FROM (
      SELECT device_id, pitch, timestamp,
             ROW_NUMBER() OVER (PARTITION BY device_id ORDER BY timestamp DESC) as rn
      FROM realtime_sensor_data
      WHERE device_id LIKE 'D%'
    ) t
    WHERE rn <= 10
    ORDER BY device_id, timestamp DESC
  `;
  
  pool.query(query, (err, rows) => {
    if (err) {
      console.error('Error fetching last 10 z_axis per Hauler:', err);
      return res.status(500).json({ error: "Error fetching z_axis values" });
    }

    const haulerData = {};
    rows.forEach(row => {
      const equipment = row.device_id;
      if (!haulerData[equipment]) {
        haulerData[equipment] = [];
      }
      haulerData[equipment].push({
        z_axis: Number(row.z_axis),
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
const getAllDevicesShiftData = (req, res) => {
  const { shift, region_id } = req.query;
  
  if (!shift || !region_id) {
    return res.status(400).json({ error: "shift and region_id required" });
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

  db.query(query, params, (err, results) => {
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
      shift,
      date: today,
      total_records: filteredResults.length,
      data: filteredResults
    });
  });
};

// ==================== DAILY (24hr - full day) ====================
// Get daily data for a specific device
/*const getDeviceDailyData = (req, res) => {
  const { device_id, region_id } = req.query;
  
  if (!device_id || !region_id) {
    return res.status(400).json({ error: "device_id and region_id required" });
  }

  const today = new Date().toISOString().split('T')[0];

  const query = `
    SELECT * FROM realtime_sensor_data 
    WHERE device_id = ?
    AND region_id = ?
    AND DATE(timestamp) = ?
    ORDER BY timestamp ASC
  `;

  db.query(query, [device_id, region_id, today], (err, results) => {
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
      date: today,
      total_records: filteredResults.length,
      data: filteredResults
    });
  });
};*/


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

// Get daily data for ALL devices in a region
/*const getAllDevicesDailyData = (req, res) => {
  const { region_id } = req.query;
  
  if (!region_id) {
    return res.status(400).json({ error: "region_id required" });
  }

  const today = new Date().toISOString().split('T')[0];

  const query = `
    SELECT * FROM realtime_sensor_data 
    WHERE region_id = ?
    AND DATE(timestamp) = ?
    ORDER BY device_id, timestamp ASC
  `;

  db.query(query, [region_id, today], (err, results) => {
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
      date: today,
      total_records: filteredResults.length,
      data: filteredResults
    });
  });
};*/


const getAllDevicesDailyData = (req, res) => {
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


// This function will be called by your existing routes
const wrapWithAnalysis = (req, res, dataFetcher) => {
  // Store the original res.json to intercept the data
  const originalJson = res.json;
  
  // Override res.json to capture the data before sending
  res.json = function(data) {
    // Check if this is the data we want to analyze
    if (data && data.data && Array.isArray(data.data)) {
      const deviceId = req.query.device_id || 'unknown';
      const timeRange = req.query.shift ? 'shift' : 
                       (req.query.date ? 'daily' : 'monthly');
      
      // Format data for Python
      const pythonInput = {
        data: data.data.map(row => ({
          time: row.timestamp || row.time,
          lat: parseFloat(row.latitude || row.lat || 0),
          lon: parseFloat(row.longitude || row.lon || 0),
          alt: parseFloat(row.altitude || row.alt || 0),
          rl: parseFloat(row.rl || 0),
          pitch: parseFloat(row.pitch || 0),
          dist: parseFloat(row.distance || row.dist || 0),
          fuel: parseFloat(row.fuel || 0),
          cost: parseFloat(row.fuel_cost || row.cost || 0)
        })),
        device_id: deviceId,
        time_range: timeRange
      };
      
      // Call Python for analysis
      const pythonProcess = exec('python routes/analysis.py', (error, stdout, stderr) => {
        if (error) {
          console.error('Python error:', error);
          // Fall back to original data if analysis fails
          return originalJson.call(res, data);
        }
        
        if (stderr) {
          console.log('Python log:', stderr);
        }
        
        try {
          const result = JSON.parse(stdout);
          
          // Decode and send DOC file
          const docBuffer = Buffer.from(result.report, 'base64');
          
          res.setHeader('Content-Type', 'application/msword');
          res.setHeader('Content-Disposition', 
            `attachment; filename=analysis_${deviceId}_${timeRange}_${Date.now()}.doc`);
          
          res.send(docBuffer);
          
        } catch (e) {
          console.error('Failed to parse Python output:', e);
          // Fall back to original data
          originalJson.call(res, data);
        }
      });
      
      pythonProcess.stdin.write(JSON.stringify(pythonInput));
      pythonProcess.stdin.end();
      
    } else {
      // Not the data we want to analyze, send normally
      originalJson.call(res, data);
    }
  };
  
  // Call the original data fetcher
  dataFetcher(req, res);
};

// ==================== ANALYSIS ENDPOINT ====================

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
      let shiftParam = '';
      if (shift === '6am-2pm') shiftParam = 'morning';
      else if (shift === '2pm-10pm') shiftParam = 'afternoon';
      else if (shift === '10pm-6am') shiftParam = 'night';
      
      req.query.shift = shiftParam;
      if (region_id) req.query.region_id = region_id;
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

  // Override res.json to capture the data
  const originalJson = res.json;
  
  res.json = function(data) {
    // Check if we have data
    if (data && data.data && Array.isArray(data.data) && data.data.length > 0) {
      const records = data.data;
      
      console.log(`✅ Fetched ${records.length} records for analysis`);
      
      // Format data for Python
      const pythonInput = {
        data: records.map(row => ({
          device_id: row.device_id || device_id,
          time: row.timestamp,
          lat: parseFloat(row.latitude || 0),
          lon: parseFloat(row.longitude || 0),
          alt: parseFloat(row.altitude || 0),
          rl: parseFloat(row.rl || 0),
          pitch: parseFloat(row.pitch || 0),
          dist: parseFloat(row.distance || 0),
          fuel: parseFloat(row.fuel || 0),
          cost: parseFloat(row.fuel_cost || 0)
        })),
        device_id: device_id,
        time_range: timeRange,
        shift: shift || null,
        is_all_devices: isAllDevices
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
          
          if (result.analysis && result.analysis.error) {
            return originalJson.call(res, { error: result.analysis.error });
          }
          
          // Decode and send DOC file
          const docBuffer = Buffer.from(result.report, 'base64');
          
          // Use filename from Python if available, otherwise generate one
          let filename = result.filename;
          if (!filename) {
            // Fallback filename generation
            const deviceName = isAllDevices ? 'AllDevices' : device_id;
            const shiftSuffix = shift ? `_${shift.replace(/[^a-zA-Z0-9]/g, '-')}` : '';
            const regionSuffix = region_id ? `_region${region_id}` : '';
            filename = `analysis_${deviceName}${regionSuffix}_${timeRange}${shiftSuffix}_${Date.now()}.docx`;
          }
          
          res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
          res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
          res.send(docBuffer);
          
          console.log(`✅ Analysis complete! Report sent: ${filename}`);
          
        } catch (e) {
          console.error('❌ Failed to parse Python output:', e);
          console.log('Raw output:', stdout.substring(0, 200));
          originalJson.call(res, { error: "Failed to generate report" });
        }
      });
      
      pythonProcess.stdin.write(JSON.stringify(pythonInput));
      pythonProcess.stdin.end();
      
    } else {
      console.log('⚠️ No data found for the selected period');
      
      // Even when no data, we still want to generate a "No Data" report
      // Format data for Python with empty array
      const pythonInput = {
        data: [],  // Empty data array
        device_id: device_id,
        time_range: timeRange,
        shift: shift || null,
        is_all_devices: isAllDevices
      };
      
      console.log(`🚀 Sending request to Python for NO DATA report...`);
      
      // Call Python script even with no data
      const pythonProcess = exec('python routes/analysis.py', (error, stdout, stderr) => {
        if (error) {
          console.error('❌ Python error:', error);
          // If Python fails, return a simple error
          return res.status(404).json({ error: "No data found for the selected period" });
        }
        
        if (stderr) {
          console.log('📝 Python log:', stderr);
        }
        
        try {
          const result = JSON.parse(stdout);
          
          // Decode and send DOC file (even if it's a "no data" report)
          const docBuffer = Buffer.from(result.report, 'base64');
          
          // Use filename from Python
          const filename = result.filename || `NoData_${isAllDevices ? 'AllDevices' : device_id}_${timeRange}${shift ? '_' + shift : ''}_${new Date().toISOString().split('T')[0]}.docx`;
          
          res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
          res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
          res.send(docBuffer);
          
          console.log(`✅ No data report sent: ${filename}`);
          
        } catch (e) {
          console.error('❌ Failed to parse Python output:', e);
          originalJson.call(res, { error: "No data found for the selected period" });
        }
      });
      
      pythonProcess.stdin.write(JSON.stringify(pythonInput));
      pythonProcess.stdin.end();
    }
  };
  
  // Call the appropriate data fetcher
  dataFetcher(req, res);
};*/

// ==================== ANALYSIS ENDPOINT ====================

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
      let shiftParam = '';
      if (shift === '6am-2pm') shiftParam = 'morning';
      else if (shift === '2pm-10pm') shiftParam = 'afternoon';
      else if (shift === '10pm-6am') shiftParam = 'night';
      
      req.query.shift = shiftParam;
      if (region_id) req.query.region_id = region_id;
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

  // Override res.json to capture the data
  const originalJson = res.json;
  
  res.json = function(data) {
    // Check if we have data
    if (data && data.data && Array.isArray(data.data) && data.data.length > 0) {
      const records = data.data;
      
      console.log(`✅ Fetched ${records.length} records for analysis`);
      
      // Format data for Python
      const pythonInput = {
        data: records.map(row => ({
          device_id: row.device_id || device_id,
          time: row.timestamp,
          lat: parseFloat(row.latitude || 0),
          lon: parseFloat(row.longitude || 0),
          alt: parseFloat(row.altitude || 0),
          rl: parseFloat(row.rl || 0),
          pitch: parseFloat(row.pitch || 0),
          dist: parseFloat(row.distance || 0),
          fuel: parseFloat(row.fuel || 0),
          cost: parseFloat(row.fuel_cost || 0)
        })),
        device_id: device_id,
        time_range: timeRange,
        shift: shift || null,
        is_all_devices: isAllDevices
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
          
          // Check if we have the report data
          if (!result.report) {
            console.error('❌ No report data in Python output');
            return originalJson.call(res, { error: "No report data generated" });
          }
          
          // Decode the base64 Excel file
          const excelBuffer = Buffer.from(result.report, 'base64');
          
          // Use filename from Python if available, otherwise generate one
          let filename = result.filename;
          if (!filename) {
            // Fallback filename generation
            const deviceName = isAllDevices ? 'AllDevices' : device_id;
            const shiftSuffix = shift ? `_${shift.replace(/[^a-zA-Z0-9]/g, '-')}` : '';
            const regionSuffix = region_id ? `_region${region_id}` : '';
            const dateStr = new Date().toISOString().split('T')[0];
            filename = `analysis_${deviceName}${regionSuffix}_${timeRange}${shiftSuffix}_${dateStr}.xlsx`;
          }
          
          // Ensure filename ends with .xlsx
          if (!filename.toLowerCase().endsWith('.xlsx')) {
            filename += '.xlsx';
          }
          
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
          console.log('Raw output:', stdout.substring(0, 500));
          
          // Try to extract error message if available
          let errorMsg = "Failed to generate report";
          try {
            const errorResult = JSON.parse(stdout);
            if (errorResult.error) {
              errorMsg = errorResult.error;
            }
          } catch (parseError) {
            // Ignore parsing error
          }
          
          originalJson.call(res, { error: errorMsg });
        }
      });
      
      pythonProcess.stdin.write(JSON.stringify(pythonInput));
      pythonProcess.stdin.end();
      
    } else {
      console.log('⚠️ No data found for the selected period');
      
      // Even when no data, we still want to generate a "No Data" report
      // Format data for Python with empty array
      const pythonInput = {
        data: [],  // Empty data array
        device_id: device_id,
        time_range: timeRange,
        shift: shift || null,
        is_all_devices: isAllDevices
      };
      
      console.log(`🚀 Sending request to Python for NO DATA report...`);
      
      // Call Python script even with no data
      const pythonProcess = exec('python routes/analysis.py', (error, stdout, stderr) => {
        if (error) {
          console.error('❌ Python error:', error);
          // If Python fails, return a simple error
          return res.status(404).json({ error: "No data found for the selected period" });
        }
        
        if (stderr) {
          console.log('📝 Python log:', stderr);
        }
        
        try {
          const result = JSON.parse(stdout);
          
          // Check if we have report data
          if (!result.report) {
            console.log('⚠️ No report data generated for empty dataset');
            return originalJson.call(res, { 
              error: "No data available for the selected period",
              details: "The system could not generate a report because no data was found."
            });
          }
          
          // Decode and send Excel file (even if it's a "no data" report)
          const excelBuffer = Buffer.from(result.report, 'base64');
          
          // Generate filename for no data report
          const deviceName = isAllDevices ? 'AllDevices' : device_id;
          const shiftSuffix = shift ? `_${shift.replace(/[^a-zA-Z0-9]/g, '-')}` : '';
          const dateStr = new Date().toISOString().split('T')[0];
          const filename = result.filename || `NoData_${deviceName}_${timeRange}${shiftSuffix}_${dateStr}.xlsx`;
          
          // Set correct headers for Excel file
          res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
          res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
          res.setHeader('Content-Length', excelBuffer.length);
          res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
          res.setHeader('Pragma', 'no-cache');
          res.setHeader('Expires', '0');
          
          res.send(excelBuffer);
          
          console.log(`✅ No data Excel report sent: ${filename}`);
          console.log(`📊 Report size: ${(excelBuffer.length / 1024).toFixed(2)} KB`);
          
        } catch (e) {
          console.error('❌ Failed to parse Python output:', e);
          
          // Return a simple JSON error as fallback
          res.status(404).json({ 
            error: "No data found for the selected period",
            details: "The system could not generate a report."
          });
        }
      });
      
      pythonProcess.stdin.write(JSON.stringify(pythonInput));
      pythonProcess.stdin.end();
    }
  };
  
  // Call the appropriate data fetcher
  dataFetcher(req, res);
};*/


// ==================== ANALYSIS ENDPOINT ====================

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
      let shiftParam = '';
      if (shift === '6am-2pm') shiftParam = 'morning';
      else if (shift === '2pm-10pm') shiftParam = 'afternoon';
      else if (shift === '10pm-6am') shiftParam = 'night';
      
      req.query.shift = shiftParam;
      if (region_id) req.query.region_id = region_id;
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

  // Override res.json to capture the data
  const originalJson = res.json;
  
  res.json = function(data) {
    // Check if we have data
    if (data && data.data && Array.isArray(data.data) && data.data.length > 0) {
      const records = data.data;
      
      console.log(`✅ Fetched ${records.length} records for analysis`);
      
      // Format data for Python - MATCHING YOUR EXISTING analysis.py EXPECTATIONS
      const pythonInput = {
        data: records.map(row => ({
          device_id: row.device_id || device_id,
          time: row.timestamp,
          lat: parseFloat(row.latitude || 0),
          lon: parseFloat(row.longitude || 0),
          pitch: parseFloat(row.pitch || 0),
          fuel: parseFloat(row.fuel || 0),
          speed: parseFloat(row.speed || 0)
        }))
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
          console.log(`📊 Devices discovered: ${result.devices_discovered?.join(', ') || 'none'}`);
          
        } catch (e) {
          console.error('❌ Failed to parse Python output:', e);
          console.log('Raw output (first 500 chars):', stdout.substring(0, 500));
          originalJson.call(res, { error: "Failed to generate report - invalid response from analysis engine" });
        }
      });
      
      pythonProcess.stdin.write(JSON.stringify(pythonInput));
      pythonProcess.stdin.end();
      
    } else {
      console.log('⚠️ No data found for the selected period');
      
      // Even when no data, send to Python for a "No Data" report
      const pythonInput = {
        data: []  // Empty data array
      };
      
      console.log(`🚀 Sending request to Python for NO DATA report...`);
      
      const pythonProcess = exec('python routes/analysis.py', (error, stdout, stderr) => {
        if (error) {
          console.error('❌ Python error:', error);
          return res.status(404).json({ error: "No data found for the selected period" });
        }
        
        if (stderr) {
          console.log('📝 Python log:', stderr);
        }
        
        try {
          const result = JSON.parse(stdout);
          
          if (result.status === 'error') {
            return res.status(404).json({ error: "No data found for the selected period" });
          }
          
          // Check if we have report data
          if (!result.report) {
            console.log('⚠️ No report data generated for empty dataset');
            return res.status(404).json({ 
              error: "No data available for the selected period"
            });
          }
          
          // Decode and send Excel file
          const excelBuffer = Buffer.from(result.report, 'base64');
          const filename = result.filename || `NoData_${device_id}_${timeRange}_${new Date().toISOString().split('T')[0]}.xlsx`;
          
          res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
          res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
          res.setHeader('Content-Length', excelBuffer.length);
          res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
          res.setHeader('Pragma', 'no-cache');
          res.setHeader('Expires', '0');
          
          res.send(excelBuffer);
          
          console.log(`✅ No data Excel report sent: ${filename}`);
          
        } catch (e) {
          console.error('❌ Failed to parse Python output:', e);
          res.status(404).json({ 
            error: "No data found for the selected period"
          });
        }
      });
      
      pythonProcess.stdin.write(JSON.stringify(pythonInput));
      pythonProcess.stdin.end();
    }
  };
  
  // Call the appropriate data fetcher
  dataFetcher(req, res);
};*/
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


/*//this is working with delay insted of region name use region_id
const fetchDashboardDataby = (req, res) => {
  const { region_id, device_id } = req.query;
  
  if (!region_id) {
    return res.status(400).json({ error: 'Region ID is required' });
  }

  console.log("Fetching for region:", region_id, "device:", device_id);

  // Helper function to remove null/undefined values
  const removeNullValues = (obj) => {
    return Object.keys(obj).reduce((acc, key) => {
      if (obj[key] !== null && obj[key] !== undefined && obj[key] !== '') {
        acc[key] = obj[key];
      }
      return acc;
    }, {});
  };

  // Step 1: Check if region exists in regions table
  const regionQuery = `SELECT * FROM regions WHERE region_id = ?`;
  
  db.query(regionQuery, [region_id], (err, regionResults) => {
    if (err) {
      console.error("Region query error:", err);
      return res.status(500).json({ error: 'Database error' });
    }
    
    if (regionResults.length === 0) {
      return res.status(404).json({ error: 'Region not found' });
    }

    const regionInfo = regionResults[0];
    
    // Step 2: Get devices for this region
    let deviceQuery = `SELECT device_id FROM devices WHERE region_id = ?`;
    const deviceParams = [region_id];

    if (device_id) {
      deviceQuery += ` AND UPPER(device_id) = UPPER(?)`;
      deviceParams.push(device_id);
    }

    db.query(deviceQuery, deviceParams, (err, devices) => {
      if (err) {
        console.error("Device query error:", err);
        return res.status(500).json({ error: 'Database error' });
      }
      
      const deviceIds = devices.map(d => d.device_id);
      
      if (deviceIds.length === 0) {
        return res.json({
          status: 'success',
          message: 'No devices found for this region',
          region: regionInfo,
          devices: [],
          data: []
        });
      }

      // Step 3: Try to get data from realtime_sensor_data
      const placeholders = deviceIds.map(() => '?').join(',');
      
      const realtimeQuery = `
        SELECT * FROM realtime_sensor_data 
        WHERE device_id IN (${placeholders})
        AND region_id = ?
        ORDER BY timestamp DESC
      `;

      db.query(realtimeQuery, [...deviceIds, region_id], (err, realtimeData) => {
        if (err) {
          console.error("Realtime query error:", err);
          return res.status(500).json({ error: 'Database error' });
        }

        // If we have realtime data, return it
        if (realtimeData.length > 0) {
          const filteredData = realtimeData.map(row => removeNullValues(row));
          
          return res.json({
            status: 'success',
            region: regionInfo,
            devices: deviceIds,
            total_records: filteredData.length,
            data: filteredData,
            source: 'realtime'
          });
        }

        // No realtime data - try dummy table
        const dummyQuery = `
          SELECT * FROM dummy 
          WHERE device_id IN (${placeholders})
          ORDER BY timestamp DESC
          LIMIT 100
        `;

        db.query(dummyQuery, deviceIds, (err, dummyData) => {
          if (err) {
            console.error("Dummy query error:", err);
            return res.status(500).json({ error: 'Database error' });
          }

          const filteredDummyData = dummyData.map(row => removeNullValues(row));

          return res.json({
            status: 'success',
            region: regionInfo,
            devices: deviceIds,
            total_records: filteredDummyData.length,
            data: filteredDummyData,
            source: 'dummy',
            message: 'No realtime data available, showing dummy data'
          });
        });
      });
    });
  });
};
*/

//i am using this
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

//before
/*const fetchDashboardDataby = (req, res) => {
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

              // ✅ If realtime data exists → return it with 'data' field
              if (hasRealtimeData && results.length > 0) {
                return res.json({
                  status: 'success',
                  source: 'realtime_sensor_data',
                  company,
                  region: regionName,
                  devices: results,
                  data: results  // ← ADD THIS LINE - same as devices
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
                  data: dummyResults  // ← Already have this
                });
              });
            }
          }
        );
      });
    }
  );
};*/

/*const fetchDashboardDataby = (req, res) => {
  const { company, region } = req.query;

  if (!company || !region)
    return res.status(400).json({ error: 'Company and region are required' });

  const regionName = region.trim();

  // Calculate 24-hour window from 6 AM today to 6 AM tomorrow
  const now = new Date();
  const startTime = new Date(now);
  startTime.setHours(6, 0, 0, 0);  // Today 6:00 AM
  
  const endTime = new Date(now);
  endTime.setDate(endTime.getDate() + 1);
  endTime.setHours(6, 0, 0, 0);  // Tomorrow 6:00 AM

  const startDateTime = startTime.toISOString().slice(0, 19).replace('T', ' ');
  const endDateTime = endTime.toISOString().slice(0, 19).replace('T', ' ');

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

      // 2️⃣ Fetch latest realtime data for each device using region_id (with time filter)
      devices.forEach((device) => {
        db.query(
          `
          SELECT *
          FROM realtime_sensor_data
          WHERE device_id = ? 
            AND region_id = ?
            AND timestamp >= ? 
            AND timestamp < ?
          ORDER BY timestamp DESC
          LIMIT 1
          `,
          [device.device_id, device.region_id, startDateTime, endDateTime],
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
};*/

// ==================== ANALYSIS ENDPOINT ====================
const generateAnalysisReport = (req, res) => {
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
      let shiftParam = '';
      if (shift === '6am-2pm') shiftParam = 'morning';
      else if (shift === '2pm-10pm') shiftParam = 'afternoon';
      else if (shift === '10pm-6am') shiftParam = 'night';
      
      req.query.shift = shiftParam;
      if (region_id) req.query.region_id = region_id;
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
  res.json = function(data) {
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
        fuel: parseFloat(row.fuel || 0),
        speed: parseFloat(row.speed || 0),
        distance: parseFloat(row.distance || 0),
        fuel_cost: parseFloat(row.fuel_cost || 0)
      }))
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
};
// ==================== EXPORT ALL FUNCTIONS ====================
module.exports = {
  register,
  signin,
  forgotPassword,
 
  insertRealtimeData,  // ✅ UPDATED: Calculates and stores ALL fields
  getLast10ZAxis,     //fuel and gradient analysis chart in kacha 
  registerToken,
  fetchDashboardData,  // ✅ UPDATED: Returns calculated fields
  
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