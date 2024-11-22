import pandas as pd
from faker import Faker
import random
import numpy as np

fake = Faker()

# Project 2: Sports Event Management - Run for Health 5K

def create_marathon_data(num_participants=1000):
    data = []
    for _ in range(num_participants):
        participant = {
            "Name": fake.name(),
            "Age": random.randint(18, 80),
            "Gender": random.choice(["Male", "Female", "Non-binary"]),
            "Email": fake.email() if random.random() > 0.1 else np.nan,  # 10% chance of having a null value
            "Phone Number": fake.phone_number(),
            "Address": fake.address(),
            "Registration Date": fake.date_this_year(),
            "Confirmed": random.choice([True, False]),
            "VIP Status": random.choice(["Regular", "VIP", "Sponsor"])
        }
        data.append(participant)

    df = pd.DataFrame(data)
    df.index += 1  # Set index to start from 1, which could be used as participant numbers
    return df

def create_marathon_times_data(num_participants=1000):
    data = []
    for i in range(1, num_participants + 1):
        participant = {
            "Participant Number": i,
            "Age": random.randint(18, 80),
            "Completion Time (minutes)": round(random.uniform(20, 120), 2),
            "Bib Number": random.randint(1000, 9999),
            "Marathon Category": random.choice(["5K", "10K", "Half Marathon", "Full Marathon"])
        }
        data.append(participant)

    df = pd.DataFrame(data)
    return df

if __name__ == "__main__":
    # Create datasets for each project
    marathon_df = create_marathon_data()
    marathon_times_df = create_marathon_times_data(len(marathon_df))

    # Save datasets to CSV files
    marathon_df.to_csv("marathon_participants.csv", index_label="Participant Number")
    marathon_times_df.to_csv("marathon_times.csv", index=False)

    print("Marathon Participants DataFrame:")
    print(marathon_df.head())
    print("\nMarathon Times DataFrame:")
    print(marathon_times_df.head())



