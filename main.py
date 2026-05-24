import pandas as pd
from pulp import *


def read_input_excel(file_name):

    df = pd.read_excel(
        file_name,
        sheet_name="Results",
        header=None
    )

    jobs = []

    col = 1

    while col < df.shape[1]:

        job_id = df.iloc[0, col]

        if pd.isna(job_id):
            break

        job = {
            "j": int(df.iloc[0, col]),
            "t": float(df.iloc[1, col]),
            "d": float(df.iloc[2, col]),
            "v": float(df.iloc[3, col]),
            "w": float(df.iloc[4, col]),
        }

        jobs.append(job)

        col += 1

    return jobs


def solve_fixed_order_lp(jobs):

    model = LpProblem(
        "Scheduling_Problem",
        LpMinimize
    )

    n = len(jobs)

    # Variables
    C = {}
    E = {}
    T = {}

    for j in range(n):

        job_num = jobs[j]["j"]

        C[job_num] = LpVariable(
            f"C_{job_num}",
            lowBound=0
        )

        E[job_num] = LpVariable(
            f"E_{job_num}",
            lowBound=0
        )

        T[job_num] = LpVariable(
            f"T_{job_num}",
            lowBound=0
        )

    # Objective Function
    model += lpSum(
        jobs[j]["v"] * E[jobs[j]["j"]] +
        jobs[j]["w"] * T[jobs[j]["j"]]
        for j in range(n)
    )

    # First job constraint
    first_job = jobs[0]

    model += (
        C[first_job["j"]] >= first_job["t"]
    )

    # Order constraints
    for j in range(1, n):

        current_job = jobs[j]
        previous_job = jobs[j - 1]

        model += (
            C[current_job["j"]]
            >=
            C[previous_job["j"]]
            + current_job["t"]
        )

    # Earliness & Tardiness constraints
    for job in jobs:

        j = job["j"]

        model += (
            E[j]
            >=
            job["d"] - C[j]
        )

        model += (
            T[j]
            >=
            C[j] - job["d"]
        )

    # Solve
    model.solve()

    print("\nSolver Status:")
    print(LpStatus[model.status])

    print("\nObjective Value:")
    print(value(model.objective))

    print("\nSolution:\n")

    for job in jobs:

        j = job["j"]

        print(
            f"Job {j} | "
            f"C={value(C[j]):.2f} | "
            f"E={value(E[j]):.2f} | "
            f"T={value(T[j]):.2f}"
        )


if __name__ == "__main__":

    jobs = read_input_excel(
        "Input_Data_File_1_Example (1).xlsx"
    )

    solve_fixed_order_lp(jobs)  