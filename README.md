# Sura Logistics

This repository is configured for **Spec-Driven Development (SDD)** powered by [OpenSpec](https://github.com/Fission-AI/OpenSpec).

---

## 🛠 Spec-Driven Development Workflow (`/opsx`)

The `/opsx` commands provide a structured, spec-first development cycle ensuring requirements, architecture, and task plans are vetted before implementation:

```
+----------------+      +----------------+      +----------------+      +----------------+
| /opsx:explore  | ---> | /opsx:propose  | ---> |  /opsx:apply   | ---> | /opsx:archive  |
| (Think & map)  |      | (Spec & plan)  |      |  (Implement)   |      | (Merge & seal) |
+----------------+      +----------------+      +----------------+      +----------------+
```

### Slash Commands

| Command | Alias | Description |
| :--- | :--- | :--- |
| `/opsx:explore` | `/opsx-explore` | Enter exploration mode to research the problem space, compare options, and sketch architectural designs without writing code. |
| `/opsx:propose <name>` | `/opsx-propose <name>` | Scaffold a change proposal and generate all spec artifacts (`proposal.md`, `design.md`, delta specs, and `tasks.md`). |
| `/opsx:apply` | `/opsx-apply` | Implement the tasks sequentially from `tasks.md` in the current active change. |
| `/opsx:sync` | `/opsx-sync` | Sync and update main specifications from delta specs when changes evolve. |
| `/opsx:archive` | `/opsx-archive` | Archive completed changes into `openspec/changes/archive/` and update the main specs. |

---

## 📁 OpenSpec Directory Structure

```
sura-logistics/
├── .agent/                  # Antigravity agent configuration
│   ├── skills/              # OpenSpec skills (propose, explore, apply, sync, archive)
│   └── workflows/           # Slash command workflow triggers (/opsx:*)
├── .agents/ -> .agent       # Compatibility symlink for Antigravity workspace
└── openspec/
    ├── specs/               # Authoritative system specifications (source of truth)
    └── changes/             # Active delta changes and proposals
        └── archive/         # Completed and archived changes
```

---

## 💻 OpenSpec CLI Reference

You can also interact directly with the `openspec` CLI:

- **List active changes:**
  ```bash
  openspec list
  ```
- **List main specs:**
  ```bash
  openspec list --specs
  ```
- **Inspect change progress:**
  ```bash
  openspec status --change "<change-name>"
  ```
- **Validate specs and changes:**
  ```bash
  openspec validate
  ```
- **Create a change manually:**
  ```bash
  openspec new change "<change-name>"
  ```
