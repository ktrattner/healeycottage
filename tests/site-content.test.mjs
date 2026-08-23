import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const rooms = JSON.parse(
  await readFile(path.join(projectRoot, "src/content/rooms.json"), "utf8")
);

test("room slugs and names are unique", () => {
  assert.equal(new Set(rooms.map(({ slug }) => slug)).size, rooms.length);
  assert.equal(new Set(rooms.map(({ name }) => name)).size, rooms.length);
});

test("room records contain valid capacity, details, and calendars", () => {
  for (const room of rooms) {
    assert.match(room.slug, /^[a-z0-9-]+$/);
    assert.ok(room.sleeps > 0, `${room.name} must sleep at least one guest`);
    assert.ok(room.details.length > 0, `${room.name} must include room details`);

    const calendarUrl = new URL(room.calendarEmbedUrl);
    assert.equal(calendarUrl.protocol, "https:");
    assert.equal(calendarUrl.hostname, "calendar.google.com");
  }
});

test("every referenced room photo exists in public output sources", () => {
  for (const room of rooms) {
    for (const photo of room.photos) {
      if (!photo.src) continue;
      const assetPath = path.join(projectRoot, "public", photo.src.replace(/^\//, ""));
      assert.ok(existsSync(assetPath), `Missing image for ${room.name}: ${photo.src}`);
      assert.ok(photo.alt, `Missing alt text for ${room.name}: ${photo.src}`);
    }
  }
});
