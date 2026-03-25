# Project Timeline Generator
A Python utility that turns a structured Excel tracker into a shareable HTML timeline for stakeholder visibility.

## Overview
`Project Timeline Generator` was built to solve a recurring communication problem during a live training rebuild: stakeholders needed a clearer, current view of deliverables, dates, and status without needing a rewritten update every time something changed.
The script reads a structured Excel tracker, applies project status logic, and generates a Gantt-style HTML timeline with a legend that can be shared from a consistent file location.

## What it does
- Opens a small GUI to create a new project or reopen a saved one
- Remembers project state and last-used project through JSON config files
- Reads a structured Excel tracker
- Generates a shareable HTML timeline using Plotly
- Color-codes project statuses
- Automatically marks incomplete past-due items as **Overdue**
- Writes the output HTML alongside the tracker file
- Opens the generated file automatically after creation
- Supports GUI and headless execution modes
- Works across Windows, macOS, and Linux/Unix

## Example tracker

This repo includes a sanitized example tracker at:

`examples/example-project-tracker.xlsx`

It is designed to demonstrate the expected structure without exposing any client-specific information.

### Required columns
The script requires these columns:
- `Task`
- `Start Date`
- `End Date`
- `Status`
If any of these are missing, the script will stop and report the missing fields.

### Optional columns
The script also supports these optional columns:
- `Task Order`
- `Sort Key`
- `Is Parent`
- `DisplayLabel`
- `Parent Slide Deck`

These are not required to generate a timeline, but they improve structure, ordering, and display quality.

### Status values
Expected status values include:
- `Complete`
- `In Progress`
- `Pending`
- `Upcoming`

The script will automatically override an item to `Overdue` when:
- the current date is past `End Date`, and
- the item is not marked `Complete`

### Example structure
The sample tracker demonstrates a simple phased project with:
- parent rows for major project phases
- child rows for individual deliverables
- task ordering and grouping fields
- realistic status progression across planning, build, review, and launch

## Installation
Install dependencies with:
```bash
python3 -m pip install -r requirements.txt
````
Or simply run the script and let it install any missing Python dependencies at runtime.

### Tkinter note
The GUI depends on `tkinter`, which is usually included with standard Python installs. It is not typically installed through `pip`.
On some Linux distributions, you may need to install it separately, for example:
```bash
sudo apt install python3-tk
```

## Usage
### GUI mode
macOS / Linux:
```bash
python3 Project_Timeline_Generator.py
```

Windows:
```bash
py Project_Timeline_Generator.py
```

### Headless mode
```bash
python3 Project_Timeline_Generator.py \
  --headless \
  --excel "./examples/example-project-tracker.xlsx" \
  --title "Example Training Renewal Project" \
  --start-date "2025-01-06" \
  --num-weeks 18
```

### Options
* `--headless` Run without GUI prompts
* `--excel` Path to the Excel tracker file
* `--config` Optional path to a JSON config file
* `--title` Override the project title
* `--start-date` Override the project start date in `YYYY-MM-DD`
* `--num-weeks` Override the number of timeline weeks to display
* `--no-open` Generate the HTML without automatically opening it

## Run the demo
### GUI mode
Run the script and select the example tracker when prompted:
macOS / Linux:
```bash
python3 Project_Timeline_Generator.py
````

Windows:
```
py Project_Timeline_Generator.py
```

### Headless mode
To run directly against the included example tracker:
```bash
python3 Project_Timeline_Generator.py \
  --headless \
  --excel "./examples/example-project-tracker.xlsx" \
  --title "Example Training Renewal Project" \
  --start-date "2030-01-07" \
  --num-weeks 10
```

### What the demo generates
The demo creates a shareable HTML timeline from the example tracker, including:
* status-based color coding
* phase and task visualization
* overdue logic for incomplete past-due items
* a built-in legend
* an output file written alongside the tracker

## Public example assets
This repo uses sanitized example data for demonstration purposes.
The included sample tracker and config file are intended to show how the tool works without exposing real stakeholder names, deliverables, timelines, or client reporting artifacts.

## Why it exists
In delivery-heavy environments, communication overhead grows quickly when timelines shift often.
This tool reduces that friction by turning structured project data into a visual artifact that stakeholders can open anytime. It was built to make project status easier to understand, easier to share upward, and easier to maintain without rewriting the same update over and over.

## Notes
* The script uses Plotly's CDN when generating the HTML output
* Saved project configuration is stored locally in `~/.gantt_generator`
* The current overdue logic uses today's date, which is ideal for active projects
* A future enhancement may add an archived/completed mode that freezes the timeline view for finished projects

## Portfolio Case Study
A full portfolio case study for this project lives here:
[Stakeholder Timeline Generator Case Study](https://portfolio.slamminstam.com/projects/stakeholder-timeline-generator/)

## License
MIT
