require('dotenv').config();
const express  = require('express');
const mongoose = require('mongoose');
const path     = require('path');
const cors     = require('cors');

const app = express();

app.use(cors({
  origin: [
    'https://master.d2pv1qhyx4ve8v.amplifyapp.com',
    'http://localhost:4000',
    'http://localhost:3000',
  ],
  credentials: true,
}));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// Routes
app.use('/api/sites',   require('./routes/sites'));
app.use('/api/upload',  require('./routes/upload'));
app.use('/api/reports', require('./routes/reports'));

// Serve dashboard for any other route
app.get('/{*splat}', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const PORT     = process.env.PORT     || 3000;
const MONGO_URI = process.env.MONGO_URI || 'mongodb://localhost:27017/sitereport';

mongoose.connect(MONGO_URI)
  .then(() => {
    console.log('MongoDB connected');
    app.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}`));
  })
  .catch(err => {
    console.error('MongoDB connection failed:', err.message);
    process.exit(1);
  });
