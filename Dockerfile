FROM node:16

WORKDIR /usr/src/app
COPY package*.json ./

RUN npm install

COPY *.ts tsconfig.json ./
RUN npm run-script build
COPY config.json credentials.json token.json ./

CMD ["node", "app.js"]
