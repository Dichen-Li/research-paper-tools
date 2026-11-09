# Research Paper Tools

Private ChatGPT plugin marketplace containing two skills:

- `paper-redraw`: reconstruct plots, diagrams, and tables from published figures.
- `evidence-first-research-slides`: create evidence-led academic presentations for Chinese-speaking CS students.

## Import into a ChatGPT workspace

1. Push this repository to a private GitHub repository.
2. In ChatGPT, open **Admin → Plugins → Add → Import marketplace**.
3. Enter the repository URL. Leave **Path** blank because the marketplace manifest is at the repository root.
4. Select the branch to follow, or leave it blank to follow the default branch.
5. Authorize GitHub, import the marketplace, and review the import result.
6. Open `research-paper-tools` and set its installation policy to **Available** or **Installed** for the intended roles.

The marketplace manifest is `.agents/plugins/marketplace.json`. The plugin itself is under `plugins/research-paper-tools/`.
