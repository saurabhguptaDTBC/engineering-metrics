# AUTOMATION - ENGINEERING METRICS


## What Is This?

This is a simple Python project intented to pull engineering metrics from multiple sources. It currently only pulls data from TargetProcess but can easily pull from other sources e.g. Azure Pipelines

____

## How To Use This

1. Fill in the relevant information in the config.py [inside src folder : Move outside source]
    1. Update the TP Access Code in config.py ( TODO : Move to environment variable) as the environment variables 
2. Run pip3 install -r requirements.txt to install dependencies
3. Run python3 populateAccelerateMetrics.py
   # For capturing Kanban Metrics : Run python3 populateAccelerateMetrics.py kanban 
        