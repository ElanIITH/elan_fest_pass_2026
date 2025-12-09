FROM node:18-alpine

WORKDIR /app
COPY package*.json ./
RUN npm install --only=production
COPY . .

ENV NODE_ENV=production

CMD ["npm", "start"]
