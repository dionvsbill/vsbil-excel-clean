// exc/server/index.js
import { createServer } from 'http';
import { app } from './app.js';
import excelRoutes from './routes/excel.js';
import roleRoutes from './routes/roles.js';
import paymentRoutes from './routes/payments.js'; // NEW: Paystack routes
import { Server } from 'socket.io';

// Health check route
app.get('/health', (req, res) => {
  res.json({ ok: true, ts: new Date().toISOString() });
});

// Mount feature routes
app.use('/excel', excelRoutes);
app.use('/roles', roleRoutes);
app.use('/payments', paymentRoutes); // NEW

// Create HTTP server
const port = process.env.PORT || 4000;
const httpServer = createServer(app);

// Attach Socket.IO for realtime collaboration
const io = new Server(httpServer, {
  cors: {
    origin: process.env.FRONTEND_URL || '*', // tighten to your frontend domain in production
    methods: ['GET', 'POST'],
  },
});

io.on('connection', (socket) => {
  console.log('Client connected:', socket.id);

  // Join a sheet room
  socket.on('join', ({ room }) => {
    socket.join(room);
    console.log(`${socket.id} joined ${room}`);
    io.to(room).emit('presence', { userId: socket.id, status: 'joined' });
  });

  // Leave a sheet room
  socket.on('leave', ({ room }) => {
    socket.leave(room);
    console.log(`${socket.id} left ${room}`);
    io.to(room).emit('presence', { userId: socket.id, status: 'left' });
  });

  // Broadcast cell edits
  socket.on('cell-edit', (op) => {
    const { room, rowIndex, colIndex, value } = op;
    socket.to(room).emit('cell-edit', { rowIndex, colIndex, value });
  });

  // Handle disconnect
  socket.on('disconnect', () => {
    console.log('Client disconnected:', socket.id);
  });
});

// Start server
httpServer.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
