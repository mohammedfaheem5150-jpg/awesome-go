#!/usr/bin/env node
'use strict';

// Validates the commands shipped by the plugins listed in
// .claude-plugin/marketplace.json. For every plugin:
//   - each commands/<name>.md file must have frontmatter with a description
//   - commands and skills must have parity: every skill has a matching
//     command of the same name, and every command has a matching skill
//   - descriptions must stay in sync: a command's frontmatter description
//     must equal its skill's frontmatter description

const fs = require('fs');
const path = require('path');
const { loadPluginDirs, parseFrontmatter } = require('./validate-skills');

const ROOT = path.resolve(__dirname, '..');

function readDescriptions(dir, filePathFor, errors) {
  const descriptions = new Map();
  if (!fs.existsSync(dir)) return descriptions;

  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const name = entry.isDirectory() ? entry.name : entry.name.replace(/\.md$/, '');
    const file = filePathFor(entry);
    if (!file) continue;
    const relFile = path.relative(ROOT, file);
    if (!fs.existsSync(file)) continue;

    const parsed = parseFrontmatter(fs.readFileSync(file, 'utf8'), relFile, errors);
    if (!parsed) continue;
    if (!parsed.fields.description || !parsed.fields.description.trim()) {
      errors.push(`${relFile}: frontmatter is missing required field "description"`);
      continue;
    }
    descriptions.set(name, { description: parsed.fields.description, file: relFile });
  }
  return descriptions;
}

function validatePlugin(pluginDir, errors) {
  const relPlugin = path.relative(ROOT, pluginDir);
  const commandsDir = path.join(pluginDir, 'commands');
  const skillsDir = path.join(pluginDir, 'skills');

  const commands = readDescriptions(
    commandsDir,
    (entry) =>
      entry.isFile() && entry.name.endsWith('.md')
        ? path.join(commandsDir, entry.name)
        : null,
    errors
  );
  const skills = readDescriptions(
    skillsDir,
    (entry) =>
      entry.isDirectory() ? path.join(skillsDir, entry.name, 'SKILL.md') : null,
    errors
  );

  for (const [name, skill] of skills) {
    const command = commands.get(name);
    if (!command) {
      errors.push(
        `${relPlugin}: skill "${name}" has no matching command file commands/${name}.md`
      );
      continue;
    }
    if (command.description !== skill.description) {
      errors.push(
        `${command.file}: description is out of sync with ${skill.file}\n` +
          `      command: ${command.description}\n` +
          `      skill:   ${skill.description}`
      );
    }
  }

  for (const name of commands.keys()) {
    if (!skills.has(name)) {
      errors.push(
        `${relPlugin}: command "${name}" has no matching skill directory skills/${name}/`
      );
    }
  }

  return commands.size;
}

function main() {
  const errors = [];
  let commandCount = 0;

  for (const pluginDir of loadPluginDirs()) {
    commandCount += validatePlugin(pluginDir, errors);
  }

  if (errors.length > 0) {
    console.error(`Command validation failed with ${errors.length} error(s):`);
    for (const error of errors) console.error(`  - ${error}`);
    process.exit(1);
  }

  console.log(`Validated ${commandCount} command(s) successfully.`);
}

main();
