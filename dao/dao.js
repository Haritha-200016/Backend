const mysql = require('mysql');
const config = require('../config/dbconfig'); // Import database config

// Create a MySQL connection pool
const pool = mysql.createPool({
    connectionLimit: 100, // You can adjust this number
    host: config.host,
    user: config.user,
    password: config.password,
    database: config.database,
    port: config.port
});

// Test the pool connection
pool.getConnection((err, connection) => {
    if (err) {
        console.error("Database connection pool failed:", err.stack);
        return;
    }
    console.log("Connected to MySQL Database with thread ID:", connection.threadId);
    connection.release(); // Always release the connection back to pool
});

module.exports = pool;