import os
import sys

# Standardize current working directory relative to project root
if os.path.basename(os.getcwd()) == 'code':
    os.chdir('..')

# Define target scripts located within the code directory
step_1_script = os.path.join('code', 'buy_process.py')
step_2_script = os.path.join('code', 'matcher.py')
step_3_script = os.path.join('code', 'sales.py')

print("Executing Step 1: Merging procurement details...")
exit_code_1 = os.system(f"{sys.executable} {step_1_script}")

if exit_code_1 == 0:
    print("\nExecuting Step 2: Running matching logic and updating warehouse...")
    exit_code_2 = os.system(f"{sys.executable} {step_2_script}")
    
    if exit_code_2 == 0:
        print("\nExecuting Step 3: Running incremental calculations and exporting sales reports...")
        # Execute the sales reconciliation script
        exit_code_3 = os.system(f"{sys.executable} {step_3_script}")
        
        if exit_code_3 == 0:
            print("\nWorkflow completed successfully! All steps executed without errors.")
        else:
            print("\nExecution failed during Step 3.")
    else:
        print("\nExecution failed during Step 2. Step 3 aborted.")
else:
    print("\nExecution failed during Step 1. Subsequent steps aborted.")