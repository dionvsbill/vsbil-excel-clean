// exc/server/realtime.js
import { createServer } from 'http';
import express from 'express';
import { Server } from 'socket.io';

const app = express();
const httpServer = createServer(app);

// Socket.IO server
const io = new Server(httpServer, {
  cors: {
    origin: '*', // tighten this to your frontend domain in production
    methods: ['GET', 'POST'],
  },
});

// Connection handler
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
    // Relay to everyone else in the room
    socket.to(room).emit('cell-edit', { rowIndex, colIndex, value });
  });

  socket.on('disconnect', () => {
    console.log('Client disconnected:', socket.id);
  });
});

// Health check
app.get('/health', (req, res) => {
  res.json({ ok: true, ts: new Date().toISOString() });
});

// Start server
const PORT = process.env.SOCKET_PORT || 4000;
httpServer.listen(PORT, () => {
  console.log(`Realtime server running on port ${PORT}`);
});
