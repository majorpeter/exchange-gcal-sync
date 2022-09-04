# exhange-gcal-sync

A simple utility that fetches calendar events from a _Microsoft Exchange_ server and synchronizes them with a _Google Calendar_ (periodically).

## Installation

Clone from Github and install dependencies:

```sh
git clone https://github.com/majorpeter/exchange-gcal-sync.git
cd exchange-gcal-sync
npm install
npx tsc
```

## Configuraton

Create a `config.json` in the application's folder, use the template `config.template.json` and update according to your setup.

## Running

See the help text of the application:

```sh
./app.js -h
```
