import os

def run():
    print("Executing Notebook Pipeline...")
    cmd = "jupyter nbconvert --to notebook --execute --inplace code/buy_process.ipynb"
    exit_code = os.system(cmd)
    if exit_code == 0:
        print("Pipeline finished successfully.")
    else:
        print(f"Pipeline crashed with exit code: {exit_code}")
    return exit_code

if __name__ == "__main__":
    run()
