const { app } = require('@azure/functions');

app.http('ping', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    return { body: 'pong - ה-API עובד' };
  }
});