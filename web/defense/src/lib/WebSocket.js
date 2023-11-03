import { readable } from 'svelte/store';
import { io } from 'socket.io-client';

let socket;

export const socketStore = readable(null, (set) => {

  // Estoy en modo devel
  // Uso hot.data para persistir el socket entre los hmr
  if (import.meta.hot) {
    if (!import.meta.hot.data.wsocket) {
      socket = io('http://localhost:8000');
      import.meta.hot.data.wsocket = socket;
    } else {
      socket = import.meta.hot.data.wsocket;
    }

  // Estoy en prod
  // Solo si no existe el socket, lo genero. Sino me empieza a duplicar en
  // cada require de este modulo, ya que debería estar por fuera de este método exportado
  // pero no puedo debido al modo desarrollo que lo necesito dentro
  } else if (!socket) {
    socket = io('http://localhost:8000');
  }

  socket.on('eventFromServer', (event) => {
    set(event);
  });

  socket.on('connect', () => {
    set({ cmd: 'id', data: socket.id });
  });

  return () => {};
});

// export const socketStore = readable(null, (set) => {
//     let socket;
  
//     // Development mode
//     if (import.meta.hot) {
//       if (!import.meta.hot.data.wsocket) {
//         socket = new WebSocket('ws://localhost:8000');
//         import.meta.hot.data.wsocket = socket;
//       } else {
//         socket = import.meta.hot.data.wsocket;
//       }
  
//     // Production mode
//     } else if (!socket) {
//       socket = new WebSocket('ws://localhost:8000');
//     }
  
//     socket.onmessage = (event) => {
//       set(JSON.parse(event.data));
//     };
  
//     socket.onopen = () => {
//       set({ cmd: 'id', data: socket.id });
//     };
  
//     return () => {
//       if (socket) {
//         socket.close();
//       }
//     };
//   });