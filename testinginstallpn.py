import re
import os

# Path to the file to be modified
file_path = os.path.join("src", "app", "plugins", "plugins.routing.ts")

# Read the content of the file
with open(file_path, "r") as file:
    content = file.read()

# Define the replacements
replacements = [
    (r"INSERT ROUTES HERE", "  },\nINSERT ROUTES HERE"),
    (r"INSERT ROUTES HERE", "      .then(m => m.ItemsPlanningPnModule)\nINSERT ROUTES HERE"),
    (r"INSERT ROUTES HERE", "    loadChildren: () => import('./modules/items-planning-pn/items-planning-pn.module')\nINSERT ROUTES HERE"),
    (r"INSERT ROUTES HERE", "    path: 'items-planning-pn',\nINSERT ROUTES HERE"),
    (r"INSERT ROUTES HERE", "  {\nINSERT ROUTES HERE"),
]

# Apply each replacement in sequence
for pattern, replacement in replacements:
    content = re.sub(pattern, replacement, content, count=1)

# Write the modified content back to the file
with open(file_path, "w") as file:
    file.write(content)
