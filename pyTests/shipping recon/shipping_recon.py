from datetime import datetime, timedelta

def get_day_suffix(day):
    if 4 <= day <= 20 or 24 <= day <= 30:
        return "th"
    else:
        return ["st", "nd", "rd"][day % 10 - 1]

def format_date(date):
    day = date.day
    suffix = get_day_suffix(day)
    month = date.strftime("%b")
    return f"{day}{suffix} {month}"

def calculate_detention_days(dis_date, ret_date, free_days, container_size):
    # Stage allocation
    stages = {
        1: 7,  # Stage 1: 7 days
        2: 3,  # Stage 2: 3 days
        3: 6,  # Stage 3: 6 days
        4: 6,  # Stage 4: 6 days
        5: None  # Stage 5: Remaining days
    }

    # Rates based on container size
    rates = {
        20: {
            2: 10150,
            3: 13485,
            4: 16820,
            5: 20155
        },
        40: {
            2: 15225,
            3: 21170,
            4: 25375,
            5: 29580
        }
    }

    # Calculate total detention days (DND-DAYS)
    dis_date = datetime.strptime(dis_date, "%d/%m/%Y")
    ret_date = datetime.strptime(ret_date, "%d/%m/%Y")
    total_days = (ret_date - dis_date).days + 1  # Including both dates

    # Initialize detention days, amounts, and rates for each stage
    detention_days = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0}
    detention_amounts = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0}
    detention_rates = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0}
    detention_ranges = {1: "", 2: "", 3: "", 4: "", 5: ""}
    days_left = total_days
    current_date = dis_date

    # Calculate detention for each stage
    for stage, stage_days in stages.items():
        if stage_days is None:  # Stage 5
            detention_days[stage] = max(0, days_left)
            if detention_days[stage] > 0:
                start_date = current_date
                end_date = current_date + timedelta(days=detention_days[stage] - 1)
                detention_ranges[stage] = f"{format_date(start_date)} - {format_date(end_date)}"
        else:
            if free_days >= stage_days:
                # Free days cover this stage
                detention_days[stage] = 0
                free_days -= stage_days  # Deduct allocated free days
            else:
                # Some detention days in this stage
                if free_days > 0:
                    detention_days[stage] = stage_days - free_days
                    free_days = 0  # All free days used
                else:
                    detention_days[stage] = min(stage_days, days_left)
                if detention_days[stage] > 0:
                    start_date = current_date
                    end_date = current_date + timedelta(days=detention_days[stage] - 1)
                    detention_ranges[stage] = f"{format_date(start_date)} - {format_date(end_date)}"
            days_left -= stage_days  # Deduct allocated stage days
            current_date += timedelta(days=stage_days)

        # Calculate detention amount and rate for stages 2 to 5
        if stage in rates[container_size]:
            detention_rates[stage] = rates[container_size][stage]
            detention_amounts[stage] = detention_days[stage] * detention_rates[stage]

    # Calculate total detention, VAT, and final detention
    total_detention = sum(detention_amounts.values())
    vat_amount = total_detention * 0.075
    final_detention = total_detention + vat_amount

    # Display results
    for stage in range(1, 6):
        print(f"Stage {stage}: {detention_days[stage]} days, Rate: {detention_rates[stage]}, Amount: {detention_amounts[stage]}, Range: {detention_ranges[stage]}")

    print(f"Total Detention: {total_detention}")
    print(f"VAT Amount (7.5%): {vat_amount}")
    print(f"Final Detention: {final_detention}")

    return final_detention

# Input from user
num_containers = int(input("Enter the number of containers (1 or 2): "))

total_final_detention = 0

for i in range(1, num_containers + 1):
    print(f"\nContainer {i}:")
    dis_date = input(f"Enter DIS-DATE for Container {i} (e.g., 01/01/2024): ")
    ret_date = input(f"Enter RET-DATE for Container {i} (e.g., 30/01/2024): ")
    free_days = int(input(f"Enter number of free days (S1-FREE) for Container {i}: "))
    container_size = int(input(f"Enter container size for Container {i} (20 or 40): "))

    final_detention = calculate_detention_days(dis_date, ret_date, free_days, container_size)
    total_final_detention += final_detention

print(f"\nTotal Final Detention for all containers: {total_final_detention}")