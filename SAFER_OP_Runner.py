from MathModels.SAFER.SAFER_OP.SAFER_Overpressure import Step5_OverPressure_Impulse, Step6a_Adjusted_P_and_I, \
    Step6b_PES_Impact
from MathModels.SAFER.SAFER_OP.SAFER_Overpressure import Step7_Final_P_and_I, Step8_Prob_Fatality_Major_Minor_Injury
from MathModels.SAFER.SAFER_OP.SAFER_Overpressure import Step9_Prob_of_Fatality_and_Injury_Glass_and_Building_Collapse, \
    Step10_Structural_Response_Complete, Sheet_Interpolator
from MathModels.SAFER.SAFER_OP.SAFER_Overpressure import Sheet_Grabber
from MathModels.SAFER.SAFER_OP.SAFER_Overpressure import Table_Crawler
import numpy as np
import pandas as pd
import time


def Overpressure_Runner(User_Inputs, Distances):
    """
    Step 2a: User Inputs for PES
    A description of each of the user inputs is provided in the following paragraphs.
    Input data includes:
    • PES building identifier
    • PES building category
    • PES building type
    • Number of people at the PES
    • Soil type
    • Compatibility Group (CG) of explosives
    • Activity type
    • Applicable environmental factors
    • Inhabited Building Distance (IBD)
    • PES Operating Hours, PES Annual Operating Hrs

        User_Input Step 3a: User Inputs for ES
    A description of each of the user inputs is provided in the following paragraphs.
    Input data includes:
    • ES building identifier
    • ES building category
    • ES building type
    • ES roof type
    • Type of glass on the ES
    • Percentage of glass on ES
    • Floor area of ES
    • Distance between the PES and ES
    • Orientation of PES to ES
    • Barricade information
    • Number of people at the ES
    • Relationship of personnel in ES to PES (Related or Unrelated/Public)
    • Number of hours people are present at ES
    • Percentage of time people are present in the ES when explosives are present in the PES
    • Upper limit (largest number of people exposed at any time during year)
    """

    """
    Inputs to Step 5:
    • Yield of the event (lbs), W1 [from Step 4]
    • Distance between the PES and ES (ft), d [from Step 3]
    
    Outputs of Step 5:
    • Unmodified pressure (psi), P
    • Unmodified impulse (psi-ms), I
    
    """
    # Non-iterable user inputs
    Y = User_Inputs['NEWQD [lbs]']
    PESTypeAndOrientation = User_Inputs['PES Type and Orientation']
    PES_Type = User_Inputs['PES Name']
    Gp = User_Inputs['Glass Percentage']
    FAes = User_Inputs['Floor Area [ft^2] of ES']
    ES_Type = User_Inputs['ES Type']
    glassType = User_Inputs['ES Glass Type']
    ES_Roof = User_Inputs['ES Roof Type']
    # IF block for 'Open' ES cases and Interpolated Sheets
    if ES_Type == 'Open':
        ES_Type = 'Open'
    else:
        sheetBuilding, sheetRoof, sheetGlass, ES_Type, ES_Roof, glassType = Sheet_Grabber(ES_Type, ES_Roof, glassType)

        # Creates Interpolation Arrays from sheets pulled by Sheet Grabber
        interpG, ABC_G = Sheet_Interpolator(sheetGlass)
        interpB, ABC_B = Sheet_Interpolator(sheetBuilding)
        interpR, ABC_R = Sheet_Interpolator(sheetRoof)

    SheetSF = 'A-11 PI Skull Fracture'
    interpSF, ABC_SF = Sheet_Interpolator(SheetSF)
    # To flip the axis in the excel file
    # 1: erase brackets around dict
    # 2: add ,orient='index'
    # 3: add axis=1 to pd.concat

    # Parallel Results array
    OP_Results = np.zeros(len(Distances), dtype=list)
    # Runner Block
    #   Loops though all distances specified by User
    Pr_sf = 100
    glassBreakage = 100
    PDb = 100
    ES_Roof_Damaged = 100
    for i in range(len(Distances)):
        d = Distances[i]
        # Non-iterable Steps
        P_Unmod, I_Unmod, Xo, Zo = Step5_OverPressure_Impulse(d, Y)
        """
        Inputs to Step 6:
        • Yield of the event (lbs), W1 [from Step 4]
        • Equivalent NEW (lbs), W2 [from Step 4]
        • Distance between the PES and ES (ft), d [from Step 3]
        • Orientation of PES to ES [from Step 3]
        • Hazard factor, Z [from Step 5]
        • Natural log of the hazard factor, X [from Step 5]
        • PES building type [from Step 2]
        
        Outputs of Step 6:
        • Adjusted pressure (psi), P′
        • Adjusted impulse (psi-ms), I′
        • Adjusted weight (lbs), Wa
        • Adjusted scaled distance (ft), Za
        • Fraction of PES roof intact, PESintact(roof)
        • Fraction of PES front wall intact, PESintact(fw)
        • Fraction of PES side walls intact, PESintact(sw)
        • Fraction of PES rear wall intact, PESintact(rw)
        • Fraction of PES roof damaged, PESdamage(roof)
        • Fraction of PES front wall damaged, PESdamage(fw)
        • Fraction of PES side walls damaged, PESdamage(sw)
        • Fraction of PES rear wall damaged, PESdamage(rw)
        """
        Wa, Za, Xa, P_Adjusted, I_Adjusted = Step6a_Adjusted_P_and_I(PESTypeAndOrientation, Y, Xo, d, Zo, P_Unmod,
                                                                     I_Unmod)
        # IF block to adjust for which steps are needed for each 'Open' case.
        if PES_Type == 'Open' and ES_Type == 'Open' or PES_Type == 'PEMB' and ES_Type == 'Open' or \
                PES_Type == 'Hollow Clay Tile' and ES_Type == 'Open':
            """
            Inputs to Step 7:
            • ES building type [from Step 3]
            • Percentage of glass on the ES (%),GP [from Step 3]
            • Floor area of the ES (ft2), FAES [from Step 3]
            • Adjusted pressure (psi), P′ [from Step 6]
            • Adjusted impulse (psi-ms), I′ [from Step 6]
            • Adjusted weight (lbs), Wa [from Step 6]
        
            Outputs of Step 7:
            • Final pressure (psi), P′′
            • Final impulse (psi-ms), I′′
            """
            P_Final, I_Final = Step7_Final_P_and_I(ES_Type, Gp, FAes, Wa, P_Adjusted, I_Adjusted)
            """
            Inputs to Step 8:
            Condition 1: Situation with no PES or ES
            • Unmodified pressure (psi), P [from Step 5]
            • Unmodified impulse (psi-ms), I [from Step 5]
            Condition 2: Situation includes a PES but no ES
            • Adjusted pressure (psi), P′ [from Step 6]
            • Adjusted impulse (psi-ms), I′ [from Step 6]
            Condition 3: Situation includes an ES
            • Final pressure (psi), P′′ [from Step 7]
            • Final impulse (psi-ms), I′′ [from Step 7]
        
            Outputs of Step 8:
            • Probability of fatality due to overpressure effects, Pf(o)
            • Probability of major injury due to overpressure effects, Pmaji(o)
            • Probability of minor injury due to overpressure effects, Pmini(o)
            """
            # IF block to implement segmentation for interpolation. Figure 5 in safe doc.

            intCapS = Pr_sf
            Pr_Fo, Pr_Major_o, Pr_Minor_o, Pr_sf = Step8_Prob_Fatality_Major_Minor_Injury(P_Final, ES_Type, I_Final,
                                                                                          intCapS, interpSF, ABC_SF)
            Pr_fb = 0
            Pr_MajorI_b = 0
            Pr_MinorI_b = 0

            """'Unmodified Pressure': P_Unmod,
              'Unmodified Impulse': I_Unmod,
              'Xo': Xo,
              'Zo': Zo,
              'Adjusted Weight': Wa,
              'Za': Za,
              'Xa': Xa,
              'Adjusted Pressure': P_Adjusted,
              'Adjusted Impulse': I_Adjusted,
              'Final Pressure': P_Final,
              'Final Impulse': I_Final,"""
            Results_Dictionary = {'Distance': d,
                                  'Probability of Fatality Due to Overpressure': Pr_Fo * 100,
                                  'Probability of Major Injury Due to Overpressure': Pr_Major_o * 100,
                                  'Probability of Minor Injury Due to Overpressure': Pr_Minor_o * 100,
                                  'Probability of Fatality due to Structure Damage': Pr_fb,
                                  'Probability of Major Injury due to Structure Damage': Pr_MajorI_b,
                                  'Probability of Minor Injury due to Structure Damage': Pr_MinorI_b}

            # Conversion of Results dictionary to array of Dataframes
            OP_Results[i] = pd.DataFrame.from_dict([Results_Dictionary])

        elif PES_Type == 'Open' and ES_Type != 'Open' or PES_Type == 'PEMB' and ES_Type != 'Open' \
                or PES_Type == 'Hollow Clay Tile' and ES_Type != 'Open':
            """
            Inputs to Step 7:
            • ES building type [from Step 3]
            • Percentage of glass on the ES (%),GP [from Step 3]
            • Floor area of the ES (ft2), FAES [from Step 3]
            • Adjusted pressure (psi), P′ [from Step 6]
            • Adjusted impulse (psi-ms), I′ [from Step 6]
            • Adjusted weight (lbs), Wa [from Step 6]
            
            Outputs of Step 7:
            • Final pressure (psi), P′′
            • Final impulse (psi-ms), I′′
            """
            P_Final, I_Final = Step7_Final_P_and_I(ES_Type, Gp, FAes, Wa, P_Adjusted, I_Adjusted)
            """
            Inputs to Step 8:
            Condition 1: Situation with no PES or ES
            • Unmodified pressure (psi), P [from Step 5]
            • Unmodified impulse (psi-ms), I [from Step 5]
            Condition 2: Situation includes a PES but no ES
            • Adjusted pressure (psi), P′ [from Step 6]
            • Adjusted impulse (psi-ms), I′ [from Step 6]
            Condition 3: Situation includes an ES
            • Final pressure (psi), P′′ [from Step 7]
            • Final impulse (psi-ms), I′′ [from Step 7]
            
            Outputs of Step 8:
            • Probability of fatality due to overpressure effects, Pf(o)
            • Probability of major injury due to overpressure effects, Pmaji(o)
            • Probability of minor injury due to overpressure effects, Pmini(o)
            """
            # IF block to implement segmentation for interpolation. Figure 5 in safe doc.
            """
                Inputs to Step 9:
                • ES building type [from Step 3]
                • ES roof type [from Step 3]
                • Type of glass on the ES [from Step 3]
                • Distance between the PES and ES (ft), d [from Step 3]
                • Percentage of glass on the ES (%), GP [from Step 3]
                • Floor area of the ES (ft2), FAES [from Step 3]
                • Adjusted pressure (psi), P′ [from Step 6]
                • Adjusted impulse (psi-ms), I′ [from Step 6]
                • Adjusted weight (lbs), Wa [from Step 6]
                • Adjusted scaled distance (ft), Za [from Step 6]
                
                Outputs of Step 9:
                • Probability of fatality due to window breakage, Pf(g)
                • Probability of major injury due to window breakage, Pmaji(g)
                • Probability of minor due to window breakage, Pmini(g)
                • Probability of fatality due to building collapse, Pf(bc)
                • Probability of major injury due to building collapse, Pmaji(bc)
                • Probability of minor injury due to building collapse, Pmini(bc)
                • Percentage of ES roof intact, %ESwall intact
                • Percentage of ES walls intact, %ESroof intact"""

            intCapS = Pr_sf
            Pr_Fo, Pr_Major_o, Pr_Minor_o, Pr_sf = Step8_Prob_Fatality_Major_Minor_Injury(P_Final, ES_Type, I_Final,
                                                                                          intCapS, interpSF, ABC_SF)
            """
            Inputs to Step 9:
            • ES building type [from Step 3]
            • ES roof type [from Step 3]
            • Type of glass on the ES [from Step 3]
            • Distance between the PES and ES (ft), d [from Step 3]
            • Percentage of glass on the ES (%), GP [from Step 3]
            • Floor area of the ES (ft2), FAES [from Step 3]
            • Adjusted pressure (psi), P′ [from Step 6]
            • Adjusted impulse (psi-ms), I′ [from Step 6]
            • Adjusted weight (lbs), Wa [from Step 6]
            • Adjusted scaled distance (ft), Za [from Step 6]

            Outputs of Step 9:
            • Probability of fatality due to window breakage, Pf(g)
            • Probability of major injury due to window breakage, Pmaji(g)
            • Probability of minor due to window breakage, Pmini(g)
            • Probability of fatality due to building collapse, Pf(bc)
            • Probability of major injury due to building collapse, Pmaji(bc)
            • Probability of minor injury due to building collapse, Pmini(bc)
            • Percentage of ES roof intact, %ESwall intact
            • Percentage of ES walls intact, %ESroof intact
            """
            intCapG = glassBreakage
            intCapB = PDb
            intCapR = ES_Roof_Damaged
            Pr_fg, Pr_fbc, Pr_MajorI_g, Pr_MinorI_g, Pr_MajorI_bc, Pr_MinorI_bc, ES_Wall_Intact, ES_Roof_Intact, \
            PDb, glassBreakage, ES_Roof_Damaged = Step9_Prob_of_Fatality_and_Injury_Glass_and_Building_Collapse(
                ES_Type, ES_Roof, FAes, glassType, Gp,
                Wa, P_Adjusted, I_Adjusted, d, Za,
                intCapG, intCapB, intCapR, interpG,
                interpB, interpR, ABC_G, ABC_B,
                ABC_R)

            """
            Inputs to Step 10:
            • Probability of fatality due to window breakage, Pf(g) [from Step 9]
            • Probability of fatality due to building collapse, Pf(bc) [from Step 9]
            • Probability of major injury due to window breakage, Pmaji(g) [from Step 9]
            • Probability of major injury due to building collapse, Pmaji(bc) [from Step 9]
            • Probability of minor injury due to window breakage, Pmini(g) [from Step 9]
            • Probability of minor injury due to building collapse, Pmini(bc) [from Step 9]
            
            Outputs of Step 10:
            • Probability of fatality due to overall building damage, Pf(b)
            • Probability of major injury due to overall building damage, Pmaji(b)
            • Probability of minor injury due to overall building damage, Pmini(b)
            """
            Pr_fb, Pr_MajorI_b, Pr_MinorI_b = Step10_Structural_Response_Complete(Pr_fg, Pr_fbc, Pr_MajorI_g,
                                                                                  Pr_MajorI_bc,
                                                                                  Pr_MinorI_g, Pr_MinorI_bc)
            """'Unmodified Pressure': P_Unmod,
            'Unmodified Impulse': I_Unmod,
            'Xo': Xo,
            'Zo': Zo,
            'Adjusted Weight': Wa,
            'Za': Za,
            'Xa': Xa,
            'Adjusted Pressure': P_Adjusted,
            'Adjusted Impulse': I_Adjusted,
            'Final Pressure': P_Final,
            'Final Impulse': I_Final,
            'Probability of Fatality due to Glass': Pr_fg,
            'Probability of Fatality due to Building Collapse': Pr_fbc,
            'Predicted building damage': PDb,
            'Probability of Major Injury due to Glass': Pr_MajorI_g,
            'Probability of Minor Injury due to Glass': Pr_MinorI_g,
            'Probability of Major Injury due Building Collapse': Pr_MajorI_bc,
            'Probability of Minor Injury due to Building Collapse': Pr_MinorI_bc,
            'Percentage of ES Wall Intact': ES_Wall_Intact,
            'Percentage of ES Roof Intact': ES_Roof_Intact,"""
            Results_Dictionary = {'Distance': d,
                                  'Probability of Fatality Due to Overpressure': Pr_Fo * 100,
                                  'Probability of Major Injury Due to Overpressure': Pr_Major_o * 100,
                                  'Probability of Minor Injury Due to Overpressure': Pr_Minor_o * 100,
                                  'Probability of Fatality Due to Glass': Pr_fg,
                                  'Probability of Major Injury Due to Glass': Pr_MajorI_g,
                                  'Probability of Minor Injury Due to Glass': Pr_MinorI_g,
                                  'Probability of Fatality Due to Building Collapse': Pr_fbc,
                                  'Probability of Major Injury Due to Building Collapse': Pr_MajorI_bc,
                                  'Probability of Minor Injury Due to Building Collapse': Pr_MinorI_bc,
                                  'Probability of Fatality due to Structure Damage': Pr_fb,
                                  'Probability of Major Injury due to Structure Damage': Pr_MajorI_b,
                                  'Probability of Minor Injury due to Structure Damage': Pr_MinorI_b}
            # 'Probability of Major Injury due to Glass': Pr_MajorI_g,
            # 'Probability of Minor Injury due to Glass': Pr_MinorI_g,
            # 'Probability of Major Injury due Building Collapse': Pr_MajorI_bc,
            # 'Probability of Minor Injury due to Building Collapse': Pr_MinorI_bc,
            # 'Percentage of ES Wall Intact': ES_Wall_Intact,
            # 'Percentage of ES Roof Intact': ES_Roof_Intact,
            # 'Probability of Fatality due to Structure Damage': Pr_fb,
            # 'Probability of Major Injury due to Structure Damage': Pr_MajorI_b,
            # 'Probability of Minor Injury due to Structure Damage': Pr_MinorI_b
            # Conversion of Results dictionary to array of Dataframes
            OP_Results[i] = pd.DataFrame.from_dict([Results_Dictionary])

        elif PES_Type != 'Open' and ES_Type == 'Open':
            damage, intact = Step6b_PES_Impact(Y, PES_Type)
            """
            Inputs to Step 7:
            • ES building type [from Step 3]
            • Percentage of glass on the ES (%),GP [from Step 3]
            • Floor area of the ES (ft2), FAES [from Step 3]
            • Adjusted pressure (psi), P′ [from Step 6]
            • Adjusted impulse (psi-ms), I′ [from Step 6]
            • Adjusted weight (lbs), Wa [from Step 6]

            Outputs of Step 7:
            • Final pressure (psi), P′′
            • Final impulse (psi-ms), I′′
            """
            P_Final, I_Final = Step7_Final_P_and_I(ES_Type, Gp, FAes, Wa, P_Adjusted, I_Adjusted)
            """
            Inputs to Step 8:
            Condition 1: Situation with no PES or ES
            • Unmodified pressure (psi), P [from Step 5]
            • Unmodified impulse (psi-ms), I [from Step 5]
            Condition 2: Situation includes a PES but no ES
            • Adjusted pressure (psi), P′ [from Step 6]
            • Adjusted impulse (psi-ms), I′ [from Step 6]
            Condition 3: Situation includes an ES
            • Final pressure (psi), P′′ [from Step 7]
            • Final impulse (psi-ms), I′′ [from Step 7]

            Outputs of Step 8:
            • Probability of fatality due to overpressure effects, Pf(o)
            • Probability of major injury due to overpressure effects, Pmaji(o)
            • Probability of minor injury due to overpressure effects, Pmini(o)
            """
            intCapS = Pr_sf
            Pr_Fo, Pr_Major_o, Pr_Minor_o, Pr_sf = Step8_Prob_Fatality_Major_Minor_Injury(P_Final, ES_Type, I_Final,
                                                                                          intCapS, interpSF, ABC_SF)
            Pr_fb = 0
            Pr_MajorI_b = 0
            Pr_MinorI_b = 0

            """'Unmodified Pressure': P_Unmod,
            'Unmodified Impulse': I_Unmod,
            'Xo': Xo,
            'Zo': Zo,
            'Adjusted Weight': Wa,
            'Za': Za,
            'Xa': Xa,
            'Adjusted Pressure': P_Adjusted,
            'Adjusted Impulse': I_Adjusted,
            'Final Pressure': P_Final,
            'Final Impulse': I_Final,"""
            Results_Dictionary = {'Distance': d,
                                  'Probability of Fatality Due to Overpressure': Pr_Fo * 100,
                                  'Probability of Major Injury Due to Overpressure': Pr_Major_o * 100,
                                  'Probability of Minor Injury Due to Overpressure': Pr_Minor_o * 100,
                                  'Probability of Fatality due to Structure Damage': Pr_fb,
                                  'Probability of Major Injury due to Structure Damage': Pr_MajorI_b,
                                  'Probability of Minor Injury due to Structure Damage': Pr_MinorI_b}

            OP_Results[i] = pd.DataFrame.from_dict([Results_Dictionary])

        else:
            damage, intact = Step6b_PES_Impact(Y, PES_Type)
            """
            Inputs to Step 7:
            • ES building type [from Step 3]
            • Percentage of glass on the ES (%),GP [from Step 3]
            • Floor area of the ES (ft2), FAES [from Step 3]
            • Adjusted pressure (psi), P′ [from Step 6]
            • Adjusted impulse (psi-ms), I′ [from Step 6]
            • Adjusted weight (lbs), Wa [from Step 6]
    
            Outputs of Step 7:
            • Final pressure (psi), P′′
            • Final impulse (psi-ms), I′′
            """
            P_Final, I_Final = Step7_Final_P_and_I(ES_Type, Gp, FAes, Wa, P_Adjusted, I_Adjusted)
            """
            Inputs to Step 8:
            Condition 1: Situation with no PES or ES
            • Unmodified pressure (psi), P [from Step 5]
            • Unmodified impulse (psi-ms), I [from Step 5]
            Condition 2: Situation includes a PES but no ES
            • Adjusted pressure (psi), P′ [from Step 6]
            • Adjusted impulse (psi-ms), I′ [from Step 6]
            Condition 3: Situation includes an ES
            • Final pressure (psi), P′′ [from Step 7]
            • Final impulse (psi-ms), I′′ [from Step 7]
    
            Outputs of Step 8:
            • Probability of fatality due to overpressure effects, Pf(o)
            • Probability of major injury due to overpressure effects, Pmaji(o)
            • Probability of minor injury due to overpressure effects, Pmini(o)
            """
            intCapS = Pr_sf
            Pr_Fo, Pr_Major_o, Pr_Minor_o, Pr_sf = Step8_Prob_Fatality_Major_Minor_Injury(P_Final, ES_Type, I_Final,
                                                                                          intCapS, interpSF, ABC_SF)
            """
            Inputs to Step 9:
            • ES building type [from Step 3]
            • ES roof type [from Step 3]
            • Type of glass on the ES [from Step 3]
            • Distance between the PES and ES (ft), d [from Step 3]
            • Percentage of glass on the ES (%), GP [from Step 3]
            • Floor area of the ES (ft2), FAES [from Step 3]
            • Adjusted pressure (psi), P′ [from Step 6]
            • Adjusted impulse (psi-ms), I′ [from Step 6]
            • Adjusted weight (lbs), Wa [from Step 6]
            • Adjusted scaled distance (ft), Za [from Step 6]

            Outputs of Step 9:
            • Probability of fatality due to window breakage, Pf(g)
            • Probability of major injury due to window breakage, Pmaji(g)
            • Probability of minor due to window breakage, Pmini(g)
            • Probability of fatality due to building collapse, Pf(bc)
            • Probability of major injury due to building collapse, Pmaji(bc)
            • Probability of minor injury due to building collapse, Pmini(bc)
            • Percentage of ES roof intact, %ESwall intact
            • Percentage of ES walls intact, %ESroof intact
            """
            intCapG = glassBreakage
            intCapB = PDb
            intCapR = ES_Roof_Damaged
            Pr_fg, Pr_fbc, Pr_MajorI_g, Pr_MinorI_g, Pr_MajorI_bc, Pr_MinorI_bc, ES_Wall_Intact, ES_Roof_Intact, \
            PDb, glassBreakage, ES_Roof_Damaged = Step9_Prob_of_Fatality_and_Injury_Glass_and_Building_Collapse(
                ES_Type, ES_Roof, FAes, glassType, Gp,
                Wa, P_Adjusted, I_Adjusted, d, Za,
                intCapG, intCapB, intCapR, interpG,
                interpB, interpR, ABC_G, ABC_B,
                ABC_R)
            """
        Inputs to Step 10:
        • Probability of fatality due to window breakage, Pf(g) [from Step 9]
        • Probability of fatality due to building collapse, Pf(bc) [from Step 9]
        • Probability of major injury due to window breakage, Pmaji(g) [from Step 9]
        • Probability of major injury due to building collapse, Pmaji(bc) [from Step 9]
        • Probability of minor injury due to window breakage, Pmini(g) [from Step 9]
        • Probability of minor injury due to building collapse, Pmini(bc) [from Step 9]

        Outputs of Step 10:
        • Probability of fatality due to overall building damage, Pf(b)
        • Probability of major injury due to overall building damage, Pmaji(b)
        • Probability of minor injury due to overall building damage, Pmini(b)
        """
            Pr_fb, Pr_MajorI_b, Pr_MinorI_b = Step10_Structural_Response_Complete(Pr_fg, Pr_fbc, Pr_MajorI_g,
                                                                                  Pr_MajorI_bc,
                                                                                  Pr_MinorI_g, Pr_MinorI_bc)

            """'Unmodified Pressure': P_Unmod,
              'Unmodified Impulse': I_Unmod,
              'Xo': Xo,
              'Zo': Zo,
              'Adjusted Weight': Wa,
              'Za': Za,
              'Xa': Xa,
              'Adjusted Pressure': P_Adjusted,
              'Adjusted Impulse': I_Adjusted,
              'Final Pressure': P_Final,
              'Final Impulse': I_Final,                            
              'Probability of Fatality due to Glass': Pr_fg,
              'Probability of Fatality due to Building Collapse': Pr_fbc,
              'Predicted building damage': PDb,
              'Probability of Major Injury due to Glass': Pr_MajorI_g,
              'Probability of Minor Injury due to Glass': Pr_MinorI_g,
              'Probability of Major Injury due Building Collapse': Pr_MajorI_bc,
              'Probability of Minor Injury due to Building Collapse': Pr_MinorI_bc,
              'Percentage of ES Wall Intact': ES_Wall_Intact,
              'Percentage of ES Roof Intact': ES_Roof_Intact,"""
            Results_Dictionary = {'Distance': d,
                                  'Probability of Fatality Due to Overpressure': Pr_Fo * 100,
                                  'Probability of Major Injury Due to Overpressure': Pr_Major_o * 100,
                                  'Probability of Minor Injury Due to Overpressure': Pr_Minor_o * 100,
                                  'Probability of Fatality due to Structure Damage': Pr_fb,
                                  'Probability of Major Injury due to Structure Damage': Pr_MajorI_b,
                                  'Probability of Minor Injury due to Structure Damage': Pr_MinorI_b}

            OP_Results[i] = pd.DataFrame.from_dict([Results_Dictionary])

    OP_Results = pd.concat(OP_Results)
    """    with pd.ExcelWriter('Overpressure_Results.xlsx') as writer:
        pd.DataFrame.from_dict([User_Inputs]).to_excel(writer, sheet_name='Inputs')
        OP_Results.to_excel(writer, sheet_name='Outputs', float_format="%.4f")"""

    return OP_Results

    # , float_format="%.2f"


if __name__ == '__main__':
    # TEST CASES FOR SAFER
    # PES Type and Orientation Options:
    """
        'ECM - Front'
        'ECM - Side'
        'ECM - Rear'
        'HAS - Front'
        'HAS - Side (W>250lb)'
        'HAS - Side (W <= 250 lbs)'
        'HAS - Rear'
        'AGBS'
        'Operating Building'
        'Ship'
        'ISO Containers'
        'Open'
    """
    # PES Name Options:
    """
        'PEMB'
        'Hollow Clay Tile'
        'HAS'
        'Large Concrete Arch ECM'
        'Medium Concrete Arch ECM'
        'Small Concrete Arch ECM'
        'Large Steel Arch ECM'
        'Medium Steel Arch ECM'
        'Small Steel Arch ECM'
        'Large AGBS'
        'Medium AGBS'
        'Small AGBS (Square)'
        'Medium Concrete Building'
        'Small Concrete Building'
        'Ship (small)'
        'Ship (medium)'
        'Ship (large)'
        'ISO Container'
        'Open'
    """
    # ES Type Options:
    """
        'Small Reinforced Concrete (Office/Commercial)'
        'Medium Reinforced Concrete (Office/Commercial)'
        'Large Reinforced Concrete Tilt-up (Commercial)'
        'Small Reinforced Masonry (Office/Commercial)'
        'Medium Reinforced Masonry (Office/Commercial)'
        'Small Unreinforced Brick (Office/Apartment)'
        'Medium Unreinforced Masonry (Office/Apartment)'
        'Large Unreinforced Masonry (Office)'
        'Small PEMB (Office/Storage)'
        'Medium PEMB (Office/Commercial)'
        'Large PEMB (Office/Storage/Hangar)'
        'Small Wood Frame (Residence)'
        'Medium Wood Frame (Residence/Apartment)'
        'Medium Steel Stud (Office/Commercial)'
        'Modular Building/Trailer (Office/Residence/Storage)'
        'Vehicle'
        'Open'
    """
    # Glass Type Options:
    """
        'Annealed'
        'Dual Pane'
        'Tempered'
    """
    # ES Roof Options:
    """
        '4 in. Reinforced Concrete'
        '14 in. Reinforced Concrete'
        'Plywood / Wood Joists (2x10 @ 16 in.)'
        'Gypsum / Fiberboard / Steel Joist'
        'Plywood panelized (2x6 @ 24 in.)'
        '2 in. Lightweight Concrete/Steel Deck & Joists'
        'Medium Steel Panel (18 gauge)'
        'Light Steel Panel (22 gauge)'
        'Steel (automobile)'
    """
    User_Inputs1 = {'Maximum Distance from Blast [ft]': 100,
                    'NEWQD [lbs]': 945,
                    'PES Type and Orientation': 'Operating Building',
                    'PES Name': 'Small Concrete Building',
                    'Glass Percentage': 20,
                    'Floor Area [ft^2] of ES': 2500,
                    'ES Type': 21,
                    'Glass Type': 42,
                    'ES Roof': 33}

    start = time.time()
    Overpressure_Runner(User_Inputs1)
    end = time.time()
    print('Execution time was: ', end - start)

