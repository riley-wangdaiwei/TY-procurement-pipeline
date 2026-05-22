import os
import sys

# Standardize current working directory relative to project root
if os.path.basename(os.getcwd()) == 'code':
    os.chdir('..')

# Define target scripts located within the code directory
step_1_script = os.path.join('code', 'buy_process.py')
step_2_script = os.path.join('code', 'matcher.py')

print("Executing Step 1: Merging procurement details...")
# Execute the initial procurement aggregation script
exit_code_1 = os.system(f"{sys.executable} {step_1_script}")

if exit_code_1 == 0:
    print("\nExecuting Step 2: Running matching logic and updating warehouse...")
    # Execute the secondary warehouse reconciliation script
    exit_code_2 = os.system(f"{sys.executable} {step_2_script}")
    
    if exit_code_2 == 0:
        print("\nWorkflow completed successfully!")
    else:
        print("\nExecution failed during Step 2.")
else:
    print("\nExecution failed during Step 1. Subsequent steps aborted.")