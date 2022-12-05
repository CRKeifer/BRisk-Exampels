import numpy as np
import scipy.stats as st


# ----------------------------------------------------------------------------------------------------------------------
def Sheet_Grabber(ES_Type, ES_Roof, glassType):
    if ES_Type == "Open":
        return ES_Type
    else:
        ESrow = Table_Crawler('Input Identifier', 'Type', ES_Type)
        sheetBuilding = ESrow['Sheet Name'].values[0]
        ES_Type = ESrow['Type'].values[0]
        if ES_Roof == "Default":
            ES_Roof = ESrow['Default Roof Type'].values[0]
            RoofRow = Table_Crawler('Input Identifier', 'Type', ES_Roof)
            sheetRoof = RoofRow['Sheet Name'].values[0]
        else:
            RoofRow = Table_Crawler('Input Identifier', 'Type', ES_Roof)
            sheetRoof = RoofRow['Sheet Name'].values[0]
            ES_Roof = RoofRow['Type'].values[0]

        GlassRow = Table_Crawler('Input Identifier', 'Type', glassType)
        sheetGlass = GlassRow['Sheet Name'].values[0]
        glassType = GlassRow['Type'].values[0]
        return sheetBuilding, sheetRoof, sheetGlass, ES_Type, ES_Roof, glassType


# ----------------------------------------------------------------------------------------------------------------------
def Sheet_Interpolator(Sheet_Name):
    import pandas as pd
    import os
    dirname = os.path.dirname(__file__)
    spreadsheet_location = os.path.join(dirname, r'.\Safer_Tables.xlsx')

    df = pd.read_excel(spreadsheet_location, sheet_name=Sheet_Name)
    header = np.array(pd.read_excel(spreadsheet_location, sheet_name=Sheet_Name, header=None, nrows=1))

    BD = np.array(df[header[0, 0]].values)
    ABC = np.array([df[header[0, 1]].values, df[header[0, 2]].values, df[header[0, 3]].values])

    BD.sort()
    ABC.sort()

    from scipy.interpolate import PchipInterpolator
    vq = PchipInterpolator(BD, ABC, axis=1)
    nmpts = 101
    rdArr = np.linspace(0, 101, nmpts)
    interpRange = vq(rdArr)

    return interpRange, ABC


# ----------------------------------------------------------------------------------------------------------------------
def Step5_OverPressure_Impulse(d, Y):
    # STEP 5: Determine open-air Pressure, Impulse (P, I).
    #   Step 5 calculates the unmodified, or open-air pressure, P, and impulse I. Values for pressure and
    #   impulse are based on simplified Kingery-Bulmash hemispherical TNT equations (Ref 5). """

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 4: Calculate the effective hazard factor, Zo, by:
    Zo = d / (Y ** (1 / 3))  # Equ(20)

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 5: Calculate Xo as the natural log of the effective hazard factor by:
    Xo = np.log(Zo)  # Equ(21)

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 6: Calculate pressure, P, and Impulse, I:

    # coefficients A, B, C, D, and E are provided in Table A-3, Pressure Calculation
    # Coefficients, based on the range of Zo.
    if 0.5 <= Zo < 7.25:
        Condition = 'Range 1'
    elif 7.25 <= Zo < 60:
        Condition = 'Range 2'
    else:
        Condition = 'Range 3'

    mask = Table_Crawler('A-3 P Calc', 'Z Range (ft/lbs1/3)', Condition)
    A = mask['A'].values[0]
    B = mask['B'].values[0]
    C = mask['C'].values[0]
    D = mask['D'].values[0]
    E = mask['E'].values[0]

    # Calculate unmodified pressure, P, by:
    # Equ (22)
    P = np.exp(A + B * Xo + C * Xo ** 2 + D * Xo ** 3 + E * Xo ** 4)

    # coefficients A, B, C, D, and E are provided in Table A-4, Impulse Calculation
    # Coefficients, based on the range of Zo.
    if 0.5 <= Zo < 2.41:
        Condition = 'Range 1'
    elif 2.41 <= Zo < 6:
        Condition = 'Range 2'
    elif 6 <= Zo < 85:
        Condition = 'Range 3'
    else:
        Condition = 'Range 4'

    mask = Table_Crawler('A-4 I Calc', 'Z Range (ft/lbs1/3)', Condition)
    A = mask['A'].values[0]
    B = mask['B'].values[0]
    C = mask['C'].values[0]
    D = mask['D'].values[0]
    E = mask['E'].values[0]

    # Calculate unmodified impulse, I, by:
    # Equ (23)
    I = np.exp(A + (B * Xo) + (C * Xo ** 2) + (D * Xo ** 3) + (E * Xo ** 4)) * (Y ** (1 / 3))
    return P, I, Xo, Zo


# ----------------------------------------------------------------------------------------------------------------------
def Step6a_Adjusted_P_and_I(PESTypeAndOrientation, Y, Xo, d, Zo, P, I):
    # STEP 6: Adjust P, I due to PES
    #   Step 6 performs two functions in SAFER. In Step 6a, SAFER calculates the pressure, P′, and
    #   impulse, I′, outside of the PES. Section 4.2.2.1 describes this process. In Step 6b, SAFER
    #   25
    #   determines the damage to the PES by calculating percentages of the PES roof and walls that
    #   remain intact following an explosive event as described in Section 4.2.2.2.

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 1: Calculate Adjusted Weight

    # Table A-5, Adjusted Weight Coefficients, provides the
    # coefficients A, B, C, D, E, F, G, and H to be used in the following equation for Wa

    if PESTypeAndOrientation == 'Open':
        Wa = Y
        Xa = Xo
        Za = Zo
        P_Adjusted = P
        I_Adjusted = I
    else:
        Condition = PESTypeAndOrientation
        if Condition == 'ECM - Front':
            if Zo > 60:
                Wa = 0.35 * Y
            elif Zo < 1.5:
                Wa = 0.1 * Y
            else:
                mask = Table_Crawler('A-5 AW', 'PES', Condition)
                A = mask['A'].values[0]
                B = mask['B'].values[0]
                C = mask['C'].values[0]
                D = mask['D'].values[0]
                E = mask['E'].values[0]
                F = mask['F'].values[0]
                G = mask['G'].values[0]
                H = mask['H'].values[0]

                # Calculate Adjusted Weight by:
                # Equ (25)
                Wa = Y * np.exp(
                    A + B * Xo + C * Xo ** 2 + D * Xo ** 3 + E * Xo ** 4 + F * Xo ** 5 + G * Xo ** 6 + H * Xo ** 7)
        elif Condition == 'ECM - Side':
            if Zo > 60:
                Wa = 0.35 * Y
            elif Zo < 2:
                Wa = 0.13 * Y
            else:
                mask = Table_Crawler('A-5 AW', 'PES', Condition)
                A = mask['A'].values[0]
                B = mask['B'].values[0]
                C = mask['C'].values[0]
                D = mask['D'].values[0]
                E = mask['E'].values[0]
                F = mask['F'].values[0]
                G = mask['G'].values[0]
                H = mask['H'].values[0]

                # Calculate Adjusted Weight by:
                # Equ (25)
                Wa = Y * np.exp(
                    A + B * Xo + C * Xo ** 2 + D * Xo ** 3 + E * Xo ** 4 + F * Xo ** 5 + G * Xo ** 6 + H * Xo ** 7)
        elif Condition == 'ECM - Rear':
            if Zo > 60:
                Wa = 0.20 * Y
            elif Zo < 2.5:
                Wa = 0.14 * Y
            else:
                mask = Table_Crawler('A-5 AW', 'PES', Condition)
                A = mask['A'].values[0]
                B = mask['B'].values[0]
                C = mask['C'].values[0]
                D = mask['D'].values[0]
                E = mask['E'].values[0]
                F = mask['F'].values[0]
                G = mask['G'].values[0]
                H = mask['H'].values[0]

                # Calculate Adjusted Weight by:
                # Equ (25)
                Wa = Y * np.exp(
                    A + B * Xo + C * Xo ** 2 + D * Xo ** 3 + E * Xo ** 4 + F * Xo ** 5 + G * Xo ** 6 + H * Xo ** 7)
        elif Condition == 'HAS - Front':
            if Zo > 63:
                Wa = 0.68 * Y
            elif Zo < 3.5:
                Wa = 0.38 * Y
            else:
                mask = Table_Crawler('A-5 AW', 'PES', Condition)
                A = mask['A'].values[0]
                B = mask['B'].values[0]
                C = mask['C'].values[0]
                D = mask['D'].values[0]
                E = mask['E'].values[0]
                F = mask['F'].values[0]
                G = mask['G'].values[0]
                H = mask['H'].values[0]

                # Calculate Adjusted Weight by:
                # Equ (25)
                Wa = Y * np.exp(
                    A + B * Xo + C * Xo ** 2 + D * Xo ** 3 + E * Xo ** 4 + F * Xo ** 5 + G * Xo ** 6 + H * Xo ** 7)
        elif Y > 250 and Condition == 'HAS - Side (W>250lb)':
            if Zo > 100:
                Wa = 1.2 * Y
            elif Zo < 2.5:
                Wa = 0.03 * Y
            else:
                mask = Table_Crawler('A-5 AW', 'PES', Condition)
                A = mask['A'].values[0]
                B = mask['B'].values[0]
                C = mask['C'].values[0]
                D = mask['D'].values[0]
                E = mask['E'].values[0]
                F = mask['F'].values[0]
                G = mask['G'].values[0]
                H = mask['H'].values[0]

                # Calculate Adjusted Weight by:
                # Equ (25)
                Wa = Y * np.exp(
                    A + B * Xo + C * Xo ** 2 + D * Xo ** 3 + E * Xo ** 4 + F * Xo ** 5 + G * Xo ** 6 + H * Xo ** 7)
        elif Y <= 250 and Condition == 'HAS - Side (W <= 250 lbs)':
            if Zo > 100:
                Wa = 0.05 * Y
            elif Zo < 2.5:
                Wa = 0.01 * Y
            else:
                mask = Table_Crawler('A-5 AW', 'PES', Condition)
                A = mask['A'].values[0]
                B = mask['B'].values[0]
                C = mask['C'].values[0]
                D = mask['D'].values[0]
                E = mask['E'].values[0]
                F = mask['F'].values[0]
                G = mask['G'].values[0]
                H = mask['H'].values[0]

                # Calculate Adjusted Weight by:
                # Equ (25)
                Wa = Y * np.exp(
                    A + B * Xo + C * Xo ** 2 + D * Xo ** 3 + E * Xo ** 4 + F * Xo ** 5 + G * Xo ** 6 + H * Xo ** 7)
        elif Condition == 'HAS - Rear':
            if Zo > 50:
                Wa = 0.07 * Y
            elif Zo < 2.67:
                Wa = 0.07 * Y
            else:
                mask = Table_Crawler('A-5 AW', 'PES', Condition)
                A = mask['A'].values[0]
                B = mask['B'].values[0]
                C = mask['C'].values[0]
                D = mask['D'].values[0]
                E = mask['E'].values[0]
                F = mask['F'].values[0]
                G = mask['G'].values[0]
                H = mask['H'].values[0]

                # Calculate Adjusted Weight by:
                # Equ (25)
                Wa = Y * np.exp(
                    A + B * Xo + C * Xo ** 2 + D * Xo ** 3 + E * Xo ** 4 + F * Xo ** 5 + G * Xo ** 6 + H * Xo ** 7)
        elif Condition == 'AGBS':
            if Zo > 140:
                Wa = 0.85 * Y
            elif Zo < 1.15:
                Wa = 0.02 * Y
            else:
                mask = Table_Crawler('A-5 AW', 'PES', Condition)
                A = mask['A'].values[0]
                B = mask['B'].values[0]
                C = mask['C'].values[0]
                D = mask['D'].values[0]
                E = mask['E'].values[0]
                F = mask['F'].values[0]
                G = mask['G'].values[0]
                H = mask['H'].values[0]

                # Calculate Adjusted Weight by:
                # Equ (25)
                Wa = Y * np.exp(
                    A + B * Xo + C * Xo ** 2 + D * Xo ** 3 + E * Xo ** 4 + F * Xo ** 5 + G * Xo ** 6 + H * Xo ** 7)
        elif Condition == 'Operating Building':
            if Zo > 140:
                Wa = 0.85 * Y
            elif Zo < 1.15:
                Wa = 0.02 * Y
            else:
                mask = Table_Crawler('A-5 AW', 'PES', Condition)
                A = mask['A'].values[0]
                B = mask['B'].values[0]
                C = mask['C'].values[0]
                D = mask['D'].values[0]
                E = mask['E'].values[0]
                F = mask['F'].values[0]
                G = mask['G'].values[0]
                H = mask['H'].values[0]

                # Calculate Adjusted Weight by:
                # Equ (25)
                Wa = Y * np.exp(
                    A + B * Xo + C * Xo ** 2 + D * Xo ** 3 + E * Xo ** 4 + F * Xo ** 5 + G * Xo ** 6 + H * Xo ** 7)
        elif Condition == 'Ship':
            if Zo > 100:
                Wa = 1.33 * Y
            elif Zo < 7.8:
                Wa = 0.5 * Y
            else:
                mask = Table_Crawler('A-5 AW', 'PES', Condition)
                A = mask['A'].values[0]
                B = mask['B'].values[0]
                C = mask['C'].values[0]
                D = mask['D'].values[0]
                E = mask['E'].values[0]
                F = mask['F'].values[0]
                G = mask['G'].values[0]
                H = mask['H'].values[0]

                # Calculate Adjusted Weight by:
                # Equ (25)
                Wa = Y * np.exp(
                    A + B * Xo + C * Xo ** 2 + D * Xo ** 3 + E * Xo ** 4 + F * Xo ** 5 + G * Xo ** 6 + H * Xo ** 7)
        else:
            if Zo > 100:
                Wa = 0.47 * Y
            else:
                Condition = 'ISO Containers'
                mask = Table_Crawler('A-5 AW', 'PES', Condition)
                A = mask['A'].values[0]
                B = mask['B'].values[0]
                C = mask['C'].values[0]
                D = mask['D'].values[0]
                E = mask['E'].values[0]
                F = mask['F'].values[0]
                G = mask['G'].values[0]
                H = mask['H'].values[0]

                # Calculate Adjusted Weight by:
                # Equ (25)
                Wa = Y * np.exp(
                    A + B * Xo + C * Xo ** 2 + D * Xo ** 3 + E * Xo ** 4 + F * Xo ** 5 + G * Xo ** 6 + H * Xo ** 7)

        # ------------------------------------------------------------------------------------------------------------------
        # Substep 2: Calculate the adjusted hazard factor, Za, by:
        Za = d / Wa ** (1 / 3)  # Equ(26)

        # ------------------------------------------------------------------------------------------------------------------
        # Substep 3: Calculate Xa as the natural log of the adjusted hazard factor by:
        Xa = np.log(Za)  # Equ(27)

        # ------------------------------------------------------------------------------------------------------------------
        # Substep 4: Calculate Adjusted P and I

        # A, B, C, D, and E are provided in Table A-3, Pressure Calculation
        # Coefficients, based on the range of Za. For “open” PES cases, P’ is set to the P value from Step 5."""
        if 0.5 <= Za < 7.25:
            Condition = 'Range 1'
        elif 7.25 <= Za < 60:
            Condition = 'Range 2'
        else:
            Condition = 'Range 3'

        mask = Table_Crawler('A-3 P Calc', 'Z Range (ft/lbs1/3)', Condition)
        A = mask['A'].values[0]
        B = mask['B'].values[0]
        C = mask['C'].values[0]
        D = mask['D'].values[0]
        E = mask['E'].values[0]

        # Calculate adjusted pressure, P′, by:
        # Equ(28)
        P_Adjusted = np.exp(A + B * Xa + C * Xa ** 2 + D * Xa ** 3 + E * Xa ** 4)

        # A, B, C, D, and E are provided in Table A-4, Impulse Calculation
        # Coefficients, based on the range of Za. For “open” PES cases, I’ is set to the value calculated for I
        # in Step 5."""

        if 0.5 <= Za < 2.41:
            Condition = 'Range 1'
        elif 2.41 <= Za < 6:
            Condition = 'Range 2'
        elif 6 <= Za < 85:
            Condition = 'Range 3'
        else:
            Condition = 'Range 4'

        mask = Table_Crawler('A-4 I Calc', 'Z Range (ft/lbs1/3)', Condition)
        A = mask['A'].values[0]
        B = mask['B'].values[0]
        C = mask['C'].values[0]
        D = mask['D'].values[0]
        E = mask['E'].values[0]
        # Calculate adjusted Impulse, I', by:
        I_Adjusted = np.exp(A + B * Xa + C * Xa ** 2 + D * Xa ** 3 + E * Xa ** 4) * Wa ** (1 / 3)

    return Wa, Za, Xa, P_Adjusted, I_Adjusted


# ----------------------------------------------------------------------------------------------------------------------
def Step6b_PES_Impact(W1, PES_Type):
    # Step 6b: Calculate PES Intact

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 1:
    locationOfDamage = ['A-6 PES Roof Damage', 'A-7 PES F Wall Damage', 'A-8 PES S Wall Damage',
                        'A-9 PES R Wall Damage']
    damage = np.zeros(4)
    Condition = PES_Type

    for i in range(4):
        mask = Table_Crawler(locationOfDamage[i], 'PES', Condition)
        Y0 = mask['Initial Breakout Value Y0 (lbs)'].values[0]
        Y100 = mask['Total Destruction Value Y100 (lbs)'].values[0]
        b = mask['b'].values[0]
        a = 1 / (Y100 - Y0) ** b
        if W1 < Y0:
            damage[i] = 0
        elif Y0 <= W1 <= Y100:
            damage[i] = a * (W1 - Y0) ** b
        else:
            damage[i] = 1

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 2:
    intact = np.zeros(4)
    for i in range(4):
        intact[i] = 1 - damage[i]

    return damage, intact


# ----------------------------------------------------------------------------------------------------------------------
def Step7_Final_P_and_I(ES_Type, Gp, FA_es, Wa, P_Adjusted, I_Adjusted):
    if ES_Type == 'Open':
        P_Final = P_Adjusted
        I_Final = I_Adjusted
    else:
        # STEP 7: Adjust P, I (due to ES).
        #   In Step 7, SAFER calculates the final pressure, P″, and final impulse, I″, values inside the ES.
        #   Given the adjusted pressure, P′, and impulse, I′, outside of the PES as determined in Step 6,
        #   another adjustment is made to determine the pressure and impulse inside the ES. If the situation
        #   29
        #   has exposed personnel in the open, this adjustment is not made because there is no structure to
        #   reduce the pressure and impulse.

        # ------------------------------------------------------------------------------------------------------------------
        # Substep 1:

        # Calculate the vent area to building volume ratio, VAVR
        # where ES height is taken from Table A-10, Pressure and Impulse Reduction Values Due to Glass
        # Percentage.
        Condition = ES_Type
        mask = Table_Crawler('A-10 PI Glass Reduction', 'ES Name', Condition)
        ESheight = mask['Height of ES (ft)'].values[0]

        VAVR = ((2.5 + Gp) / 100) / (FA_es * ESheight)  # Equ(32)

        # ------------------------------------------------------------------------------------------------------------------
        # Substep 2:
        # Calculate the average reduction value, RVave. If the VAVR is greater than 0.005, calculate the
        # average reduction value by:
        if VAVR > 0.005:
            RVave = 0.3 * Wa ** 0.095 * P_Adjusted ** (-0.32)  # Equ(33)
        else:
            RVave = (0.3 * Wa ** 0.095 * P_Adjusted ** (-0.32)) * (0.392 + 0.0568 * np.log10(Wa))  # Equ(34)

        # ------------------------------------------------------------------------------------------------------------------
        # Substep 3
        # The RVave calculated in Substep 2 is used to complete the definition of the 3 points of the
        # pressure reduction function in Table 10, Pressure Reduction Function Parameters.

        # ------------------------------------------------------------------------------------------------------------------
        # Substep 4:
        # The reduction level, RL, is dependent on the percentage of glass that is entered by the user. The
        # entered percentage of glass is compared to the average protection (shown in Table A-10,
        # Pressure and Impulse Reduction Values Due to Glass Percentage) to determine which line
        # segment of the pressure reduction function (shown in Figure 4) is applicable. Then, using the line
        # segment defined in Table 10, Pressure Reduction Function Parameters, and the calculated
        # average reduction level, the equation of the line is determined in the form:
        # RL = (slope * Gp/100) + y-intercept"""

        maxReduction = 0.5
        averageProtection = 0.075
        fullVenting = 0.25

        point1X = 0
        point1Y = maxReduction
        point2X = averageProtection
        point2Y = RVave
        point3X = fullVenting
        point3Y = 0

        # solving for the slope for the 2 lines
        slope1_2 = (point2Y - point1Y) / (point2X - point1X)
        slope2_3 = (point3Y - point2Y) / (point3X - point2X)

        # solving for the y-intercept of point 2 and 3
        c2_3 = point3Y - (point3X * slope2_3)  # y = mx + b --- -b = mx - y --- y - mx = b
        yIntercept2_3 = c2_3

        if 0 < Gp < 7.5:
            RL = (slope1_2 * (Gp / 100)) + point1Y  # Equ(35)
        else:
            RL = (slope2_3 * (Gp / 100)) + yIntercept2_3  # Equ(35)

        # ------------------------------------------------------------------------------------------------------------------
        # Substep 5
        # This reduction level is used to calculate P″ and I″
        P_Final = (1 - RL) * P_Adjusted

        I_Final = (1 - RL) * I_Adjusted
    return P_Final, I_Final


# ----------------------------------------------------------------------------------------------------------------------
def Step8_Prob_Fatality_Major_Minor_Injury(P_Final, ES_Type, I_Final, intCap, interpSF, ABC_SF):
    # STEP 8: Assess Pf(o), Pmaji(o), Pmini(o).
    # Step 8 completes the Pressure and Impulse Branch by determining the probability of fatality,
    # Pf(o), probability of major injury, Pmaji(o), and probability of minor injury due to the effects of
    # pressure and impulse, Pmini(o). In determining the probability of fatality, major injury, and minor

    # injury, SAFER calculates three consequences:
    # -Lung Rupture
    # -Whole body displacement
    # -Skull Fracture

    # Calculations in Step 8 are grouped in five parts. Step 8a performs additional pressure and
    # impulse calculations; Step 8b determines the probability of fatality and injury from lung rupture;
    # Step 8c determines the probability of fatality and injury from whole body displacement; Step 8d
    # determines the probability of fatality and injury from skull fracture; and Step 8e aggregates all
    # probabilities of fatality and injury to determine the overall probability of fatality and injury due
    # to the effects of pressure and impulse, Pf(o), Pmaji(o), and Pmini(o)

    # There are three potential input conditions:
    # -Condition 1: Situation with no PES or ES
    #    • Unmodified pressure (psi), P [from Step 5]
    #    • Unmodified impulse (psi-ms), I [from Step 5]
    # -Condition 2: Situation includes a PES but no ES
    #    • Adjusted pressure (psi), P′ [from Step 6]
    #    • Adjusted impulse (psi-ms), I′ [from Step 6]
    # -Condition 3: Situation includes an ES
    #   • Final pressure (psi), P″ [from Step 7]
    #    • Final impulse (psi-ms), I″ [from Step 7]

    # With the values of pressure and impulse known (from Step 5, 6, or 7), the human vulnerability
    # due to direct pressure and impulse effects is calculated.
    # SAFER considers the human vulnerability due to the effects of pressure and impulse to be a
    # function of lung rupture, whole body displacement, or skull fracture (or the combination of the
    # three). The probability of fatality due to lung rupture, body displacement, or skull fracture is
    # based on the probit functions originally published by The Netherlands Organization for Applied
    # Scientific Research TNO (Ref 9). Those functions determine the probability of fatality as a
    # function of incident pressure and impulse, the ambient atmospheric pressure, and an assumed
    # mass of the human body.

    # Prior to using the probit functions to determine the human vulnerability, SAFER must first
    # calculate reflected pressure, dynamic pressure, nominal pressure, scaled pressure, and scaled
    # impulse. """

    # ------------------------------------------------------------------------------------------------------------------
    # Step 8a: Pressure and Impulse Calculations
    #   ambient pressure, Pambient is assumed to be 14.5 psi.
    P_Ambient = 14.5

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 1:
    #   Calculate reflected pressure, Preflected, by:
    # Equ(38)
    P_Reflected = (2 * P_Final * ((4 * P_Final) + (7 * P_Ambient))) / (P_Final + (7 * P_Ambient))

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 2:
    #   Calculate dynamic pressure, Pdynamic, by:
    P_Dynamic = (2.5 * P_Final ** 2) / (7 * P_Ambient + P_Final)  # Equ(39)

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 3:
    #   If the ES in the situation is open, calculate the nominal pressure, Pnominal, by:
    if ES_Type == "Open":
        P_Nominal = P_Final + P_Dynamic  # Equ(40)
    else:
        P_Nominal = P_Reflected  # Equ(41)

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 4:
    #   Calculate the scaled pressure, Pscaled, by:
    P_Scaled = (P_Dynamic * 6.895) / (P_Ambient * 0.001)  # Equ(42)

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 4:
    #   Calculate the scaled impulse, Iscaled, by:
    #   where I” is from Step 7 and the constant is based on the conversion of pressure, time, and mass
    #   to the appropriate English units
    I_Scaled = I_Final * 0.005291  # Equ(43)

    # The parameters calculated in Step 8a are used as inputs to the remainder of Step 8 equations for
    # calculating the probability of fatality or injury due to lung rupture, whole body displacement, and
    # skull fracture

    # ------------------------------------------------------------------------------------------------------------------
    # //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    # ------------------------------------------------------------------------------------------------------------------

    # Step 8b: Lung Rupture
    #   Step 8b determines the probability of fatality and major and minor injuries resulting from lung
    #   rupture. This is accomplished by calculating the s and z parameters associated with a standard
    #   normal curve used in the TNO probit functions. The TNO probit functions are based on the
    #   standard normal distribution translated by subtracting 5 from the z value. However, SAFER uses
    #   the standard normal distribution without translation.

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 1:
    #   To determine the probability of fatality due to lung rupture, Pf(lr), SAFER calculates slr by:
    S_lr = 4.2 / P_Scaled + 1.3 / I_Scaled  # Equ(44)

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 2:
    #   Using the calculated slr, SAFER calculates zlr by:
    Z_lr = -5.74 * np.log(S_lr)  # Equ(45)

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 3:
    #   Given the value calculated for zlr, SAFER determines Pf(lr) by using a normal distribution where
    #   Pf(lr) is equal to the area under the standard normal distribution to the left of the zlr value.
    Pr_Flr = st.norm.cdf(Z_lr)

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 4:
    #   Using the pressure calculated from Step 7 (P’’), calculate the probability of a major injury from
    #   lung rupture (Pmaji(lr)). The relationship between the pressure and the probability of major injury
    #   is estimated by curve-fitting actual data with the following linear function.
    Pr_Major_lr = 0.01 * P_Final - 0.18  # Equ(46)

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 5:
    #   Using the pressure calculated from Step 7 (P’’), calculate the probability of a minor injury from
    #   lung rupture (Pmini(lr)). The relationship between the pressure and the probability of minor injury
    #   is estimated by curve-fitting actual data with the following linear function.
    Pr_Minor_lr = 0.032 * P_Final - 0.046

    # ------------------------------------------------------------------------------------------------------------------
    # //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    # ------------------------------------------------------------------------------------------------------------------

    # Step 8c: Whole Body Displacement
    #   Step 8c determines the probability of fatality, major injury, and minor injury resulting from
    #    whole body displacement. As in Step 8b, this is accomplished by calculating the s and z
    #   parameters associated with a standard normal distribution and the TNO probit functions.

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 1:
    #   To determine the probability of fatality due to body displacement, Pf(bd), SAFER calculates sbd
    #   by:
    # Equ(48)
    S_bd = 7280 / (P_Nominal * 6895) + (1.3 * 10 ** 9) / ((P_Nominal * 6895) * (I_Final * 6.895))

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 2:
    #   Using the calculated sbd, SAFER calculates zbd by
    Z_bd = -2.44 * np.log(S_bd)  # Equ(49)

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 3:
    #   Given the value calculated for zbd, SAFER determines Pf(bd) by using a normal distribution where
    #   Pf(bd) is equal to the area under the standard normal distribution to the left of the zbd value.
    Pr_Fbd = st.norm.cdf(Z_bd)

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 4:
    #   Given the value for probability of fatality due to whole body displacement (Pf(bd)), SAFER
    #   determines the probability of a major injury from whole body displacement (Pmaji(bd)).
    Pr_Major_bd = 1 - np.exp(-7 * Pr_Fbd)  # Equ(50)

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 5:
    #   Given the value for probability of major injury from whole body displacement (Pmaji(bd)), SAFER
    #   determines the probability of a minor injury from whole body displacement (Pmini(bd)).
    Pr_Minor_bd = 1 - np.exp(-7 * Pr_Major_bd)  # Equ(51)

    # ------------------------------------------------------------------------------------------------------------------
    # //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    # ------------------------------------------------------------------------------------------------------------------

    # Step 8d: Skull Fracture
    #   Step 8d determines the probability of fatality, major injury, and minor injury resulting from skull
    #   fracture. As in Steps 8b and 8c, this is accomplished by calculating the s and z parameters
    #   associated with a standard normal distribution and the TNO probit functions.

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 1:
    #   To determine the probability of fatality due to skull fracture, Pf(sf), SAFER calculates ssf by:
    S_sf = 2430 / (P_Nominal * 6895) + (4 * 10 ** 8) / ((P_Nominal * 6895) * (I_Final * 6.895))

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 2:
    #   Using the calculated ssf, SAFER calculates zsf by:
    Z_sf = -8.49 * np.log(S_sf)

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 3:
    #   Given the value calculated for zsf, SAFER determines Pf(sf) by using a normal distribution where
    #   Pf(sf) is equal to the area under the standard normal distribution to the left of the zsf value.
    Pr_Fsf = st.norm.cdf(Z_sf)

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 4:
    #   To determine the probability of major injury due to skull fracture, SAFER calculates the
    #   probability of skull fracture, P(sf) using P’’ and I” from Step 7. The probability of skull fracture
    #   uses the same hyperbolae interpolation methodology from Section 4.3.1.2, Substep 1. The skull
    #   fracture hyperbolae parameters are contained in Table A-11, Pressure-impulse Coefficients for
    #   P(sf).
    # C = (P_Final - A) * (I_Final - B)
    Pr_sf = Hyperbolic_Interpolation(ABC_SF, interpSF, P_Final, I_Final, intCap) / 100

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 5:
    #   Given the value of probability of skull fracture, P(sf), the probability of major injury due to skull
    #   fracture, Pmaji(sf), is determined
    if Pr_sf < 0.01:
        Pr_Major_sf = 0.25 * Pr_sf
    else:
        Pr_Major_sf = -1.34 * Pr_sf ** 2 + 2.09 * Pr_sf + 0.25

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 5:
    #   Given the probability of major injury due to skull fracture, Pmaji(sf), calculate the probability of a
    #   minor injury from skull fracture (Pmini(sf)).
    if Pr_Major_sf < 0.01:
        Pr_Minor_sf = 10 * Pr_Major_sf
    else:
        Pr_Minor_sf = -1.34 * Pr_Major_sf ** 2 + 2.09 * Pr_Major_sf + 0.25

    # ------------------------------------------------------------------------------------------------------------------
    # //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    # ------------------------------------------------------------------------------------------------------------------

    # Step 8e: Aggregation of Consequences
    #   Given the probability of fatality for lung rupture, whole body displacement, and skull fracture,
    #   calculate the probability of fatality due to the effects of pressure and impulse, Pf(o), by:

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 1:
    #   Given the probability of fatality for lung rupture, whole body displacement, and skull fracture,
    #   calculate the probability of fatality due to the effects of pressure and impulse, Pf(o), by:
    # Equ(56)
    Pr_Fo = (Pr_Flr + ((1 - Pr_Flr) * Pr_Fbd) + ((1 - Pr_Flr) * (1 - Pr_Fbd) * Pr_Fsf))
    if Pr_Fo < 0:
        Pr_Fo = 0
    elif Pr_Fo > 1:
        Pr_Fo = 1
    else:
        Pr_Fo = Pr_Fo
    # ------------------------------------------------------------------------------------------------------------------
    # Substep 2:
    #   Calculate the probability of major injury due to the effects of pressure and impulse, Pmaji(o), by:
    # Equ(57)
    Pr_Major_o = (
            Pr_Major_lr + ((1 - Pr_Major_lr) * Pr_Major_bd) + ((1 - Pr_Major_lr) * (1 - Pr_Major_bd) * Pr_Major_sf))
    if Pr_Major_o < 0:
        Pr_Major_o = 0
    elif Pr_Major_o > 1:
        Pr_Major_o = 1
    else:
        Pr_Major_o = Pr_Major_o
    # ------------------------------------------------------------------------------------------------------------------
    # Substep 3:
    #   Calculate the probability of fatality due to the effects of pressure and impulse, Pmini(o), by:
    # Equ(58)
    Pr_Minor_o = (
            Pr_Minor_lr + ((1 - Pr_Minor_lr) * Pr_Minor_bd) + ((1 - Pr_Minor_lr) * (1 - Pr_Minor_bd) * Pr_Minor_sf))
    if Pr_Minor_o < 0:
        Pr_Minor_o = 0
    elif Pr_Minor_o > 1:
        Pr_Minor_o = 1
    else:
        Pr_Minor_o = Pr_Minor_o
    return Pr_Fo, Pr_Major_o, Pr_Minor_o, Pr_sf


# ----------------------------------------------------------------------------------------------------------------------
def Step9_Prob_of_Fatality_and_Injury_Glass_and_Building_Collapse(ES_Type, ES_Roof, FA_es, glassType, Gp, Wa,
                                                                  P_Adjusted,
                                                                  I_Adjusted, d, Za, intCapG, intCapB, intCapR,
                                                                  interpG, interpB, interpR, ABC_G, ABC_B, ABC_R):
    # STEP 9: Determine adjusted P, I effect on ES (building collapse and glass hazard).
    #   Step 9 performs three functions in SAFER. In Step 9a, SAFER calculates the probability of
    #   fatality due to window breakage, Pf(g). Section 4.3.1.1 describes the procedures for this step.
    #   In Step 9b, SAFER calculates the probability of fatality due to building collapse, Pf(bc). Section
    #   4.3.1.2 describes the procedures for this step. In Step 9c, SAFER determines the damage to the
    #   ES by calculating percentages of the ES roof and walls that remain intact following an
    #   explosives event. Section 4.3.1.3 describes the procedures for this step.

    # ------------------------------------------------------------------------------------------------------------------
    # //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    # ------------------------------------------------------------------------------------------------------------------

    # 4.3.1.1 Step 9a: Human vulnerability due to window breakage
    #   This step determines the probability of fatality due to window breakage, Pf(g), the probability of
    #   major injury due to window breakage, Pmaji(g), and the probability of minor injury due to glass
    #   breakage, Pmini(g). To determine Pf(g), SAFER calculates the probability of a person being in the
    #   glass hazard area followed by the probability of a major injury given that the person is in a glass
    #   hazard area. Finally, SAFER determines the probability of fatality based on the probability of
    #   major injury.

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 1:
    #   the aspect ratio = 2 for all ES building types except modular/trailers. The aspect ratio = 3
    #   for an ES building type of modular/trailers.

    if ES_Type == "Modular/Trailers":
        aspectRatio = 3
    else:
        aspectRatio = 2
    # To determine the probability of a person being in the glass hazard area, SAFER calculates
    # Potential Window Hazard Floor Area, PWHFA, by:
    # Equ(59)
    PWHFA = 22.5 * ((FA_es * aspectRatio) ** (1 / 2) + (FA_es / aspectRatio) ** (1 / 2))

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 2:
    #   Calculate the probability of a person being in the glass hazard area, Pgha, by:
    Pr_gha = (Gp / 100) * (PWHFA / FA_es)  # Equ(60)
    # This equation simply represents the percentage of glass present multiplied by the ratio of the
    # glass hazard area to the total area.

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 3:
    #   The pressure and impulse at the ES (P′ and I′) are used with stored Pressure-impulse diagrams to
    #   determine the percentage of glass broken, % glass breakage (Attachment 6). This method uses P
    #   and I coefficients from Table A-19, Pressure-impulse Coefficients – Glass Breakage, as
    #   described in Section 4.3.1.2, Substep 1.
    glassBreakage = Hyperbolic_Interpolation(ABC_G, interpG, P_Adjusted, I_Adjusted, intCapG)
    # ------------------------------------------------------------------------------------------------------------------
    # Substep 4:
    #   To determine the probability of a major injury given that the person is in a glass hazard area,
    #   SAFER calculates the base probability of major injury, Pbase, by:

    # where the coefficients M and N are provided in Table A-12, Power Curve Parameters for Major
    # Injury as a Function of Glass Breakage, and are based on the type of glass on the ES.
    Condition = glassType
    mask = Table_Crawler('A-12 PC Params Glass Maj Inj', 'Glass Type', Condition)
    M = mask['M'].values[0]
    N = mask['N'].values[0]

    Pr_base = M * glassBreakage ** N  # Equ(61)

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 5:
    #   The base probability of major injury, Pbase, is associated with a fixed yield; therefore, it must be
    #   adjusted to the relative NEWQD experienced at the ES. A yield adjustment factor is calculated by:

    # Coefficients A, B, and C are provided in Table A-13, Yield Adjustment Curve
    # Parameters, and are based on the type of glass on the ES.
    # Rationale for the yield adjustment curves and the associated parameters in this substep is detailed
    # in Attachment 6.
    Condition = glassType
    mask = Table_Crawler('A-13 Yield Adjust Params', 'Glass Type', Condition)
    A = mask['A'].values[0]
    B = mask['B'].values[0]
    C = mask['C'].values[0]

    Y_Actual = Wa
    Y_Nominal = 50000
    G = 100 / glassBreakage
    S = (Y_Nominal / Y_Actual) ** (1 / 3)
    R = np.log(Y_Actual / Y_Nominal)
    Y_Adjusted = (A * R) + B * G ** (C * R * S)

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 6:
    #   SAFER Version 3.1 takes higher-velocity glass fragments into account for annealed or dual pane
    #   glass when the 100% glass breakage level is met or exceeded. For annealed and dual pane glass
    #   with less than 100% glass breakage and for tempered glass regardless of percent glass breakage,
    #   calculate the probability of a major injury given that the person is in the glass hazard area, Ppha,
    #   by:

    # For annealed glass with 100% glass breakage, calculate the Ppha as follows:
    if glassBreakage == 100 and glassType == "Annealed":
        # Calculate the impulse-adjusted probability of major injury, IA, by:
        IA = (8.1216 * Y_Actual ** 0.015541) * np.log(I_Adjusted) - (18.103 * Y_Actual ** 0.066969)
        # To ensure a smooth transition region at the 100% glass breakage level, a continuity correction is
        # introduced. Calculate the maximum continuity correction, CCmax, by:
        CC_max = -0.337 + 0.26 * np.log(Y_Actual) - 0.0103 * (np.log(Y_Actual)) ** 2
        # Calculate the transition point, TP, by:
        TP = np.exp(1.1924 + 0.66148 * np.log(Y_Actual) - 0.010167 * (np.log(Y_Actual)) ** 2)
        # Calculate the actual continuity correction, CCa, by:
        CC_a = ((CC_max - 1) / TP) * d + 1
        # Calculate the probability of major injury, Ppha, by:
        Pr_pha = IA * CC_a
    # For dual pane glass with 100% glass breakage, calculate the Ppha as follows:
    elif glassBreakage == 100 and glassType == "Dual Pane":
        # Calculate the impulse-adjusted probability of major injury, IA, by:
        IA = (7.0757 * Y_Actual ** 0.035394) * np.log(I_Adjusted) - (15.233 * Y_Actual ** 00.086094)
        # Calculate the maximum continuity correction, CCmax, by:
        CC_max = -1.413 + 0.43 * np.log(Y_Actual) - 0.0171 * (np.log(Y_Actual)) ** 2
        # Calculate the transition point, TP, by:
        TP = np.exp(0.89573 + 0.73204 * np.log(Y_Actual) - 0.013288 * (np.log(Y_Actual)) ** 2)
        # Calculate the actual continuity correction, CCa, by:
        CC_a = ((CC_max - 1) / TP) * d + 1
        # Calculate the probability of major injury, Ppha, by:
        Pr_pha = IA * CC_a
    else:
        Pr_pha = Pr_base * Y_Adjusted

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 7:
    #   SAFER uses the assumption that one glass fatality occurs per 30 major injuries. Therefore,
    #   SAFER calculates the initial probability of fatality due to window breakage, Pf(gi), by:
    Pr_fgi = Pr_gha * Pr_pha * (1 / 30)

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 8:
    #   Za is the Adjusted Scaled Distance
    #   SAFER adjusts the Pf(gi) based on the scaled distance. The glass adjustment, G, is calculated by:
    G2 = 10.958 - 0.417 * Za
    # If G is greater than or equal to 24, the value is reset to 1; if G is less than 12, the value is reset to
    # 6. Then, the probability of fatality due to window breakage is calculated by:
    if G2 >= 24:
        G2 = 1
    elif G2 < 12:
        G2 = 6
    else:
        G2 = G2
    Pr_fg = G2 * Pr_fgi

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 9:
    #   If necessary, SAFER determines the probability of fatality due to glass in the “close-in” or
    #   transition region.

    # A and B are provided in Table 11, Close-in Adjustment Parameters for Glass Fatality, Wa
    # is from Step 5 or 6, GP is from Step 3, and FAES is from Step 3.
    # R1, R2, Pf(g)1 are provided in Table 11, Close-in Adjustment Parameters for Glass Fatality,
    # and are based on the type of glass on the ES. Then Pf(g)2 is calculated by:
    Condition = glassType
    mask = Table_Crawler('T-11 GF Close-in Adjust Params', 'Window Type', Condition)
    R1 = mask['R1 (ft/lb^(1/3))'].values[0]
    R2 = mask['R2 (ft/lb^(1/3))'].values[0]
    A = mask['Pfg2 A'].values[0]
    B = mask['Pfg2 B'].values[0]
    Pr_fg1 = mask['Pfg1'].values[0]

    Pr_fg2 = (A + B * np.log10(Wa)) * (Gp / 10) * (5000 / FA_es) ** (1 / 2)

    if Za < R1:
        Pr_fg_final = Pr_fg1
    elif Za >= R2:
        Pr_fg_final = Pr_fg
    else:
        Pr_fg_final = Pr_fg1 - ((Za - R1) / (R2 - R1)) ** 2 * (Pr_fg1 - Pr_fg2)

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 10:
    #   SAFER calculates the probability of major injury due to window breakage, Pmaji(g), by:
    Pr_MajorI_g = Pr_fg_final * 30

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 11:
    #   SAFER calculates the probability of minor injury due to window breakage, Pmini(g), by:
    Pr_MinorI_g = Pr_fg_final * 500

    # ------------------------------------------------------------------------------------------------------------------
    # //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    # ------------------------------------------------------------------------------------------------------------------

    # 4.3.1.2 Step 9b: Human vulnerability due to building collapse
    #   This step determines the probability of fatality due to building collapse, Pf(bc), probability of
    #   major injury due to building collapse, Pmaji(bc), and probability of minor injury due to building
    #   collapse, Pmini(bc).

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 1:
    #   ES damage is determined using standard pressure and impulse diagrams of the form shown in
    #   Figure 5 using the adjusted pressure, P′, and adjusted impulse, I′, from Step 6. Damage curves
    #   are hyperbolae, which are defined by the standard equation for a hyperbola:
    #   C = (P - A) * (I - B)
    #   There are 16 families of these curves, one for each ES (Ref 10). Table A-17, Pressure / Impulse
    #   Coefficients – ES Building Percent Damage, provides the constants used to generate each family
    #   of hyperbolae. The damage to the ES is determined in SAFER using an interpolation routine.
    #   This interpolation results in the predicted ES building damage, PDb.
    PDb = Hyperbolic_Interpolation(ABC_B, interpB, P_Adjusted, I_Adjusted, intCapB)
    # ------------------------------------------------------------------------------------------------------------------
    # Substep 2:
    #   Then, using the PDb and the truncated normal distribution curves shown in Figure 6, the
    #   probability of fatality as a function of structural damage, Pf(bc), is found. The parameters defining
    #   the truncated normal curves are shown in Table A-14, Structure Damage / Fatality Normal
    #   Distribution Parameters.
    if PDb < 30:
        Pr_fbc = 0.0001
    else:
        Condition = ES_Type
        mask = Table_Crawler('PolyfitCoeffs', 'ES Building Type', Condition)
        List = mask.values[0]
        C_Array = List[1:13]
        Pr_fbc = np.polyval(C_Array, PDb)

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 3:
    #   If appropriate, SAFER determines the probability of fatality due to building collapse in the
    #   “close-in” or transition region.

    #  A, B, C and R2min are provided in Table A-16, Close-in Adjustment Parameters for
    #  Building Collapse Region Boundaries, and Wa is from Step 5 or 6.
    Condition = ES_Type
    mask = Table_Crawler('A-16 Close-in Adjst Params BCRB', 'ES Building Type', Condition)
    R1 = mask['R1 (ft/lbs^(1/3))'].values[0]
    R2min = mask['R2 (ft/lbs^(1/3))'].values[0]
    A = mask['A'].values[0]
    B = mask['B'].values[0]
    C = mask['C'].values[0]

    R2 = A + ((B * Wa) / (C + Wa))

    if R2 > R2min:
        R2 = R2
    else:
        R2 = R2min

    # Pf(bc)1 and Pf(bc)2 are provided in Table A-15, Close-in Adjustment Parameters for Pf(bc), is
    # provided in Table A-16, Close-in Adjustment Parameters for Building Collapse Region
    # Boundaries, and R2 is calculated by:
    Condition = ES_Type
    if Za < R1:
        mask = Table_Crawler('A-15 Close-in Adjst Params Pfbc', 'ES Building Type', Condition)
        Pr_fbc_final = mask['Pf(bc)1'].values[0]

    elif Za >= R2:
        Pr_fbc_final = Pr_fbc  # from Substep 2

    else:
        mask = Table_Crawler('A-15 Close-in Adjst Params Pfbc', 'ES Building Type', Condition)
        Pr_fbc1 = mask['Pf(bc)1'].values[0]
        if ES_Type == 'Small Unreinforced Brick (Office/Apartment)' or ES_Type == 'Modular Building/Trailer (' \
                                                                                  'Office/Residence/Storage)':
            Pr_fbc2 = mask['Pf(bc)2'].values[0]
        else:
            Pr_fbc2 = eval(mask['Pf(bc)2'].values[0])
        Pr_fbc_final = Pr_fbc1 - ((Za - R1) / (R2 - R1)) ** 2 * (Pr_fbc1 - Pr_fbc2)

    Pr_fbc_final *= 100

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 4:
    #   SAFER calculates the probability of major injury due to building collapse, Pmaji(bc). First a
    #   nominal Pmaji(bc) is calculated by:

    # the PDb is from Substep 1, major injury damage offset, maj(i)DO, and major injury
    # maximum, maj(i)max, is from Table 12, Major and Minor Injury Parameters for Building
    # Collapse. MAXInjury is a constant set to 100% and MINInjury is a constant set to 0.
    Condition = ES_Type
    mask = Table_Crawler('T-12 Injury Params BC', 'ES Type', Condition)
    maj_i_DO = mask['Damage offset (maj(i)DO)'].values[0]
    maj_i_max = mask['Maximum (maj(i)max)'].values[0]
    IFR = mask['Injury Fatality Ratio (IFR)'].values[0]

    MAXinjury = 100
    MINinjury = 0

    Pr_MajorI_bc_Nominal = ((maj_i_max - MINinjury) / (MAXinjury - maj_i_DO) * PDb) - (
            (maj_i_max * maj_i_DO) / (MAXinjury - maj_i_DO))

    # An adjusted probability of major injury due to building collapse is determined by

    Pr_MajorI_bc_Adjusted = IFR * Pr_fbc_final

    # where IFR is the injury to fatality ratio from Table 12, Major and Minor Injury Parameters for
    # Building Collapse. A comparison is made between the Nom. Pmaji(bc) and Adj. Pmaji(bc). The
    # variable with the highest value is assigned as the final Pmaji(bc).
    if Pr_MajorI_bc_Nominal < Pr_MajorI_bc_Adjusted:
        Pr_MajorI_bc = Pr_MajorI_bc_Adjusted
    else:
        Pr_MajorI_bc = Pr_MajorI_bc_Nominal

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 5:
    #   Similar to the calculation of probability of major injury due to building collapse, the probability
    #   of a minor injury from building collapse (Pmini(bc)) is determined. First a nominal Pmini(bc) is
    #   calculated by:

    # PDb is from Substep 1, minor injury damage offset, min(i)DO, and minor injury plateau
    # damage, min(i)PD, is from Table 12, Major and Minor Injury Parameters for Building Collapse.
    Condition = ES_Type
    mask = Table_Crawler('T-12 Injury Params BC', 'ES Type', Condition)
    min_i_DO = mask['Damage offset (min(i)DO)'].values[0]
    min_i_PD = mask['Plateau damage (min(i)PD)'].values[0]
    IFR = mask['Injury Fatality Ratio (IFR)'].values[0]

    Pr_MinorI_bc_Nominal = ((PDb - min_i_DO) / (min_i_PD - min_i_DO)) * 100

    # An adjusted probability of major injury due to building collapse is determined by:
    Pr_MinorI_bc_Adjusted = (IFR ** 2) * Pr_fbc_final

    if Pr_MinorI_bc_Adjusted < Pr_MinorI_bc_Nominal:
        Pr_MinorI_bc = Pr_MinorI_bc_Nominal
    else:
        Pr_MinorI_bc = Pr_MinorI_bc_Adjusted

    # ------------------------------------------------------------------------------------------------------------------
    # //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    # ------------------------------------------------------------------------------------------------------------------

    # 4.3.1.3 Step 9c: ES Roof and Wall Damage

    # ------------------------------------------------------------------------------------------------------------------
    # Substep 1:
    #   The wall damage is set equal to the predicted building damage, PDb (from Section 4.3.1.2,
    #   Substep 1).
    ES_Wall_Damaged = PDb
    # ------------------------------------------------------------------------------------------------------------------
    # Substep 2:
    #   The roof damage, %ESroof damaged, uses the same hyperbolic interpolation methodology from
    #   Section 4.3.1.2, Substep 1. The roof hyperbolae parameters are contained in Table A-18,
    #   Pressure-impulse Coefficients – ES Roof Damage.
    if ES_Roof == "Gypsum / Fiberboard / Steel Joist":
        if P_Adjusted >= 0.1044 and I_Adjusted >= 1.047:
            ES_Roof_Damaged = 100
        else:
            ES_Roof_Damaged = 0
    else:
        ES_Roof_Damaged = Hyperbolic_Interpolation(ABC_R, interpR, P_Adjusted, I_Adjusted, intCapR)
    # ------------------------------------------------------------------------------------------------------------------
    # Substep 3:
    #   The percentage of the walls and roof intact is the percentage not damaged, as shown by:
    ES_Wall_Intact = 100 - ES_Wall_Damaged
    ES_Roof_Intact = 100 - ES_Roof_Damaged

    # Capping upper and lower bounds.

    if Pr_MinorI_g > 100:
        Pr_MinorI_g = 100
    elif Pr_MinorI_g < 0:
        Pr_MinorI_g = 0
    else:
        Pr_MinorI_g = Pr_MinorI_g

    if Pr_MajorI_g > 100:
        Pr_MajorI_g = 100
    elif Pr_MajorI_g < 0:
        Pr_MajorI_g = 0
    else:
        Pr_MajorI_g = Pr_MajorI_g

    if Pr_fg_final < 0:
        Pr_fg_final = 0
    elif Pr_fg_final > 100:
        Pr_fg_final = 100
    else:
        Pr_fg_final = Pr_fg_final

    if Pr_MinorI_bc > 100:
        Pr_MinorI_bc = 100
    elif Pr_MinorI_bc < 0:
        Pr_MinorI_bc = 0
    else:
        Pr_MinorI_bc = Pr_MinorI_bc

    if Pr_MajorI_bc > 100:
        Pr_MajorI_bc = 100
    elif Pr_MajorI_bc < 0:
        Pr_MajorI_bc = 0
    else:
        Pr_MajorI_bc = Pr_MajorI_bc

    if Pr_fbc_final < 0:
        Pr_fbc_final = 0
    elif Pr_fbc_final > 100:
        Pr_fbc_final = 100
    else:
        Pr_fbc_final = Pr_fbc_final

    return Pr_fg_final, Pr_fbc_final, Pr_MajorI_g, Pr_MinorI_g, Pr_MajorI_bc, Pr_MinorI_bc, ES_Wall_Intact, \
           ES_Roof_Intact, PDb, glassBreakage, ES_Roof_Damaged


# ----------------------------------------------------------------------------------------------------------------------
def Step10_Structural_Response_Complete(Pr_fg_final, Pr_fbc_final, Pr_MajorI_g, Pr_MajorI_bc, Pr_MinorI_g,
                                        Pr_MinorI_bc):
    # 4.3.2 STEP 10: Assess Pf(b), Pmaji(b), Pmini(b)
    #   Step 10 completes the Structural Response Branch by determining the probability of fatality due
    #   to overall building damage, Pf(b), probability of major injury due to overall building damage,
    #    Pmaji(b) , probability of minor due to overall building damage, Pmini(b).
    #   Inputs to Step 10:
    #   • Probability of fatality due to window breakage, Pf(g) [from Step 9]
    #   • Probability of fatality due to building collapse, Pf(bc) [from Step 9]
    #   • Probability of major injury due to window breakage, Pmaji(g) [from Step 9]
    #   • Probability of major injury due to building collapse, Pmaji(bc) [from Step 9]
    #   • Probability of minor injury due to window breakage, Pmini(g) [from Step 9]
    #   • Probability of minor injury due to building collapse, Pmini(bc) [from Step 9]
    # Sets all Probabilities to decimal form.
    Pr_fg_final /= 100
    Pr_fbc_final /= 100
    Pr_MajorI_g /= 100
    Pr_MinorI_g /= 100
    Pr_MajorI_bc /= 100
    Pr_MinorI_bc /= 100

    Pr_fb = (Pr_fg_final + ((1 - Pr_fg_final) * Pr_fbc_final)) * 100
    Pr_MajorI_b = (Pr_MajorI_g + ((1 - Pr_MajorI_g) * Pr_MajorI_bc)) * 100
    Pr_MinorI_b = (Pr_MinorI_g + ((1 - Pr_MinorI_g) * Pr_MinorI_bc)) * 100

    return Pr_fb, Pr_MajorI_b, Pr_MinorI_b


# ----------------------------------------------------------------------------------------------------------------------
def Table_Crawler(Table_Name, Column_Header, Condition):
    import pandas as pd
    import os
    dirname = os.path.dirname(__file__)
    spreadsheet_location = os.path.join(dirname, r'.\Safer_Tables.xlsx')
    df = pd.read_excel(spreadsheet_location, sheet_name=Table_Name)
    mask = df.loc[df[Column_Header] == Condition]  # Strips a row from panda's dataframe.
    return mask


# ----------------------------------------------------------------------------------------------------------------------
def Hyperbolic_Interpolation(ABC, interpRange, Pressure, Impulse, intCap):
    intCap = int(intCap) + 1
    idxP = abs(interpRange[0, :] - Pressure).argmin()
    idxI = abs(interpRange[1, :] - Impulse).argmin()

    if idxI > idxP:
        sheetCap = idxI
    else:
        sheetCap = idxP

    if intCap < sheetCap:
        intCap = intCap
    else:
        intCap = sheetCap

    Pmin = min(ABC[0, :])
    Imin = min(ABC[1, :])
    cpi = (Pressure - interpRange[0, 0:intCap]) * (Impulse - interpRange[1, 0:intCap])
    error = abs(cpi - interpRange[2, 0:intCap])
    if Pressure < Pmin or Impulse < Imin or len(error) == 0:
        percent = 0.0000001
    else:
        percent = np.argmin(error)

    if percent == 0:
        percent = 0.0000001
    else:
        percent = percent

    return percent
