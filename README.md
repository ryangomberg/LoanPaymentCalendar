# LoanPaymentCalendar
*Amortization schedule and master payment calendar using Excel VBA*

## Project Overview
This project constructs two medians of consolidating loan amortization data upon feeding information into an Excel encoded UserForm. First, a new worksheet is created under the borrower's full name, including common features of an amortization schedule (payment date, payment amount, principal, interest, outstanding balance). Second, the data is fed into a master calendar, containing all payments in chronological order, as well as monthly worksheets in which a payment is due. This opens up opportunities for cross-sectional analysis, as through a PivotTable or PivotChart. The UserForm does not accept missing entries or entries with the incorrect specifications.



## Primary Objectives
Build a UserForm that will construct an amortization schedule. The UserForm will request the following parameters: 
* Loan amount (in USD)
* Effective annual interest rate, expressed in decimal form
* Number of payments each year
* Total number of payments
* Date of the first payment
* Name of the borrower

In return, an Excel worksheet under the borrower's name will have an amortization schedule, containing the date, payment amount, principal, interest, and outstanding balance after each installment. Additionally, each payment is routed into the payment calendar workbook.

## Methods
### Submodules
OpenAmortForm(): Opens the Amortization Scheduling Form; this macro is tied to the "Build Schedule" button.

### Forms
AmortizationSchedulingForm

#### Functions


## Repository Structure

## Technologies
* Microsoft Excel
* Visual Basic Access (VBA)
* UserForms

## Skills Demonstrated
