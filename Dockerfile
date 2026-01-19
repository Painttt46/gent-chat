FROM node:20-alpine

# Install graphicsmagick and ghostscript for PDF to image conversion
RUN apk add --no-cache graphicsmagick ghostscript

WORKDIR /app

COPY package*.json ./

RUN npm ci --omit=dev

COPY . .

EXPOSE 3000

CMD ["node", "index.js"]
