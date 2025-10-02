// db.js
import pg from 'pg';

const { Pool } = pg;

const pool = new Pool({
    connectionString: process.env.POSTGRES_URL,
});

// ใช้ export default
export default {
    query: (text, params) => pool.query(text, params),
};