# MS-Project-VBA-Macros
Contains macros written for MS Project tasks.
Have written macros more PERT where optimistic, most likely, and pessimistic durations can be entered in predefined fields Duration1, 2 and 3 respectively. The weights for these durations can also be set in Number1, 2 and 3 fields respectively, which are predefined in MS Project. This weights will define how the Estimate (E) is calculated and placed in the Duration field (predefined in MS Project). E.g. if you place weight as Optimistic - 1, Most likely - 1, and Pessimistic - 4; then (E) will be calculated as 
E = [(Optimistic * 1) + (Most likely * 1) + (Pessimistic * 4)]
                              6
This macro will also calculate Standard Deviation in the Duration4 field, to calculate Sigma values to understand the over all project duration.

It also contains code for Adding, and Hiding below mentioned fields.
Field List and its description
  1. Duration1 - Optimistic Duration
  2. Duration2 - Most likely Duration
  3. Duration3 - Pessimistic Duration
  4. Duration4 - Standard Deviation
  5. Number1 - Optimistic Weight
  6. Number2 - Most likely Weight
  7. Number3 - Pessimistic Weight
