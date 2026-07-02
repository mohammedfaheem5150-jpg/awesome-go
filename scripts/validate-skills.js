#!/usr/bin/env node
'use strict';

// Validates every skill shipped by the plugins listed in
// .claude-plugin/marketplace.json. Each skill lives at
// <plugin>/skills/<skill-name>/SKILL.md and must carry YAML frontmatter
// with a `name` matching its directory and a non-empty `description`.

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const NAME_PATTERN = /^[a-z0-9]+(-[a-z0-9]+)*$/;
const MAX_NAME_LENGTH = 64;
const MAX_DESCRIPTION_LENGTH = 1024;

function loadPluginDirs() {
  const marketplacePath = path.join(ROOT, '.claude-plugin', 'marketplace.json');
  const marketplace = JSON.parse(fs.readFileSync(marketplacePath, 'utf8'));
  return (marketplace.plugins || []).map((plugin) =>
    path.resolve(ROOT, plugin.source)
  );
}

// Minimal frontmatter parser: supports `key: value` pairs with optional
// single or double quotes. Enough for SKILL.md and command frontmatter.
function parseFrontmatter(content, file, errors) {
  const match = content.match(/^---\r?\n([\s\S]*?)\r?\n---\r?\n?([\s\S]*)$/);
  if (!match) {
    errors.push(`${file}: missing YAML frontmatter block delimited by ---`);
    return null;
  }
  const fields = {};
  for (const line of match[1].split(/\r?\n/)) {
    if (!line.trim() || line.trim().startsWith('#')) continue;
    const kv = line.match(/^([A-Za-z][\w-]*):\s*(.*)$/);
    if (!kv) {
      errors.push(`${file}: unparseable frontmatter line: ${JSON.stringify(line)}`);
      continue;
    }
    let value = kv[2].trim();
    if (
      (value.startsWith('"') && value.endsWith('"') && value.length >= 2) ||
      (value.startsWith("'") && value.endsWith("'") && value.length >= 2)
    ) {
      value = value.slice(1, -1);
    }
    fields[kv[1]] = value;
  }
  return { fields, body: match[2] };
}

function validateSkill(skillDir, errors) {
  const skillName = path.basename(skillDir);
  const skillFile = path.join(skillDir, 'SKILL.md');
  const relFile = path.relative(ROOT, skillFile);

  if (!fs.existsSync(skillFile)) {
    errors.push(`${relFile}: skill directory has no SKILL.md`);
    return;
  }

  const parsed = parseFrontmatter(fs.readFileSync(skillFile, 'utf8'), relFile, errors);
  if (!parsed) return;
  const { fields, body } = parsed;

  if (!fields.name) {
    errors.push(`${relFile}: frontmatter is missing required field "name"`);
  } else {
    if (fields.name !== skillName) {
      errors.push(
        `${relFile}: frontmatter name "${fields.name}" does not match directory name "${skillName}"`
      );
    }
    if (!NAME_PATTERN.test(fields.name)) {
      errors.push(`${relFile}: name "${fields.name}" must be lowercase letters, digits, and hyphens`);
    }
    if (fields.name.length > MAX_NAME_LENGTH) {
      errors.push(`${relFile}: name exceeds ${MAX_NAME_LENGTH} characters`);
    }
  }

  if (!fields.description || !fields.description.trim()) {
    errors.push(`${relFile}: frontmatter is missing required field "description"`);
  } else if (fields.description.length > MAX_DESCRIPTION_LENGTH) {
    errors.push(`${relFile}: description exceeds ${MAX_DESCRIPTION_LENGTH} characters`);
  }

  if (!body.trim()) {
    errors.push(`${relFile}: skill body is empty`);
  }
}

function main() {
  const errors = [];
  let skillCount = 0;

  for (const pluginDir of loadPluginDirs()) {
    const skillsDir = path.join(pluginDir, 'skills');
    if (!fs.existsSync(skillsDir)) continue;

    const entries = fs
      .readdirSync(skillsDir, { withFileTypes: true })
      .filter((entry) => entry.isDirectory());

    for (const entry of entries) {
      skillCount += 1;
      validateSkill(path.join(skillsDir, entry.name), errors);
    }
  }

  if (errors.length > 0) {
    console.error(`Skill validation failed with ${errors.length} error(s):`);
    for (const error of errors) console.error(`  - ${error}`);
    process.exit(1);
  }

  console.log(`Validated ${skillCount} skill(s) successfully.`);
}

if (require.main === module) {
  main();
}

module.exports = { loadPluginDirs, parseFrontmatter };
