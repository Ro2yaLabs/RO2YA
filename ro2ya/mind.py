import pandas as pd
import os
import sqlite3

def mind(user_responses, database):
    path = database
    connection = sqlite3.connect(path)

    try:
        def read_table(table_name):
            query = f"SELECT * FROM {table_name}"
            df = pd.read_sql_query(query, connection)
            return df

        df_prs = read_table("assessment_assets_PRS")
        df_vak = read_table("assessment_assets_VAK")
        df_emq = read_table("assessment_assets_EMQ")
        df_trs = read_table("assessment_assets_TRS")
        df_ctd = read_table("assessment_assets_CTD")
        df_cmf = read_table("assessment_assets_CMF")
        df_qtm = read_table("assessment_assets_QTM")

        scoring_methods = {
        "A1": {1: 1, 2: 2, 3: 3, 4: 4},
        "A2": {1: 4, 2: 3, 3: 2, 4: 1},
        "A3": {1: 1, 2: 4, 3: 4, 4: 1},
        "A4": {1: 4, 2: 1, 3: 1, 4: 4}}

        traits = ["EI", "SN", "TF", "JP"]

        trait_results = {trait: [] for trait in traits}
        trait_letters = {}

        personality = {}

        for trait in traits:
            y_indices = set(df_prs[df_prs[trait] == "Y"]["Idx"].tolist())
            trait_responses = [response for idx, response in enumerate(user_responses) if str(float(idx + 1)) in y_indices]

            count_1_or_2, count_3_or_4 = 0, 0
            for i in trait_responses:
                if i in (1, 2):
                    count_1_or_2 += 1
                else:
                    count_3_or_4 += 1

            probability_e = count_1_or_2 / max(1, len(trait_responses))
            probability_i = count_3_or_4 / max(1, len(trait_responses))
            
            personality[trait[0]] = round(probability_e, 2)
            personality[trait[1]] = round(probability_i, 2)

            if probability_e > probability_i:
                determined_trait = "E"
            else:
                determined_trait = "I"

            trait_letter = trait[0] if determined_trait == "E" else trait[1]
            trait_letters[trait] = trait_letter

        personality_type = "".join([trait_letters[trait] for trait in traits])
        personality["title"] = personality_type
        
        v_score, a_score, ks_score, kf_score, kp_score = 0, 0, 0, 0, 0

        for idx, response in enumerate(user_responses):

            vn_value, kf_value, kp_value, as_value, ks_value = df_vak.loc[idx, ["VN", "KF", "KP", "AS", "KS"]]

            if response in [1, 2]:
                v_score += scoring_methods.get(vn_value, {}).get(response, 0)
                kf_score += scoring_methods.get(kf_value, {}).get(response, 0)
                kp_score += scoring_methods.get(kp_value, {}).get(response, 0)
            elif response in [3, 4]:
                a_score += scoring_methods.get(as_value, {}).get(response, 0)
                ks_score += scoring_methods.get(ks_value, {}).get(response, 0)

        Visual, Auditory, Kinesthetic = v_score, a_score, ks_score + kf_score + kp_score

        type = "Visual" if Auditory<Visual>Kinesthetic else "Auditory" if Visual<Auditory>Kinesthetic else "Kinesthetic"

        vak = {
            "type": type,
            "visual": Visual,
            "auditory": Auditory,
            "kinesthetic": Kinesthetic
        }
        
        scores_emq = {"SelfAwarness": 0, "ManagingEmotions": 0, "MotivatingOneself": 0, "Empathy": 0, "SocialSkills": 0}
        for idx, row in df_emq.iterrows():
            response = user_responses[idx]

            for category in scores_emq.keys():
                scoring_method = row[category]
                if pd.notna(scoring_method) and scoring_method != "":
                    scores_emq[category] += scoring_methods[scoring_method][response]


        ei = {
            "SelfAwarness": scores_emq["SelfAwarness"],
            "ManagingEmotions": scores_emq["ManagingEmotions"],
            "MotivatingOneself": scores_emq["MotivatingOneself"],
            "Empathy": scores_emq["Empathy"],
            "SocialSkills": scores_emq["SocialSkills"]
        }

        roles = {
            "RI": "Resource Investigator",
            "CO": "Co-ordinator",
            "PL": "Plant",
            "SH": "Shaper",
            "ME": "Monitor Evaluator",
            "IMP": "Implementer",
            "TW": "Teamworker",
            "CF": "Completer Finisher",
            "SP": "Specialist"
        }

        scores_trs = {role: 0 for role in roles.keys()}

        for idx, row in df_trs.iterrows():
            response = user_responses[idx]

            for role in scores_trs.keys():
                scoring_method = row[role]
                if pd.notna(scoring_method) and scoring_method != "":
                    scores_trs[role] += scoring_methods[scoring_method][response]

        rls = scores_trs

        scores_ctd = {col: 0 for col in df_ctd.columns if col != "Idx"}

        for idx, row in df_ctd.iterrows():
            response = user_responses[idx]

            for skill, scoring_method in row.items():
                if pd.notna(scoring_method) and scoring_method != "" and scoring_method in scoring_methods:
                    scores_ctd[skill] += scoring_methods[scoring_method][response]

        scores_qtm = {trait: 0 for trait in df_qtm.columns if trait != "Idx"}

        for idx, row in df_qtm.iterrows():
            response = user_responses[idx]

            for trait, scoring_method in row.items():
                if pd.notna(scoring_method) and scoring_method != "" and scoring_method in scoring_methods:
                    scores_qtm[trait] += scoring_methods[scoring_method][response]

        def determine_trait_level(score):
            if score > 15:
                return "balanced"
            elif 10 <= score <= 15:
                return "rebalanced"
            else:
                return "unbalanced"

        traits = {}
        for trait, score in scores_qtm.items():
            traits[trait] = {"score":score, "level": determine_trait_level(score)}

        # The CTD Section

        def determine_skill_level(score):
            if score >= 30:
                return "Strong Skill"
            elif 25 <= score < 30:
                return "Needs Attention"
            else:
                return "Development Priority"

        def determine_aspect_level(score, num_skills):
            max_score = 36 if num_skills == 3 else 48
            percentage = (score / max_score) * 100
            if percentage >= 81:
                return "Advanced", percentage
            elif 58 <= percentage < 81:
                return "Intermediate", percentage
            else:
                return "Beginner", percentage

        def determine_space_level(score, num_aspects):
            max_score = {3: 108, 4: 144, 5: 180}
            percentage = (score / max_score[num_aspects]) * 100
            if percentage >= 72:
                return "Advanced", percentage
            elif 54 <= percentage < 72:
                return "Intermediate", percentage
            else:
                return "Beginner", percentage
            
        hierarchical_mapping = {}
        for _, row in df_cmf.iterrows():
            space, aspect, skill = row
            if space not in hierarchical_mapping:
                hierarchical_mapping[space] = {}
            if aspect not in hierarchical_mapping[space]:
                hierarchical_mapping[space][aspect] = []
            hierarchical_mapping[space][aspect].append(skill)
        hierarchical_mapping.pop('Space', None)

        scores = {col: 0 for col in df_ctd.columns if col != "Idx"}
        for idx, row in df_ctd.iterrows():
            response = user_responses[idx]
            for skill, scoring_method in row.items():
                if pd.isna(scoring_method) or scoring_method not in scoring_methods:
                    continue
                scores[skill] += scoring_methods[scoring_method][response]

        # Aggregate scores for each aspect and space
        space_scores = {}
        aspect_scores = {}
        skill_scores = {}
        for space, aspects in hierarchical_mapping.items():
            space_score = 0
            space_data = {}
            for aspect, skills_list in aspects.items():
                aspect_score = sum(scores[skill] for skill in skills_list if skill in scores)
                space_score += aspect_score
                level, percentage = determine_aspect_level(aspect_score, len(skills_list))
                skills_data = {skill: (scores[skill], determine_skill_level(scores[skill])) for skill in skills_list if
                            skill in scores}
                space_data[aspect] = {"Score": aspect_score, "Level": level, "Percentage": percentage, "Skills": skills_data}
                aspect_scores[aspect] = {"level": level, "percentage": round(percentage, 2)}
                for skill, skill_data in skills_data.items():
                    skill_scores[skill] = {"score": skill_data[0], "level": skill_data[1]}
            level, percentage = determine_space_level(space_score, len(aspects))
            space_scores[space] = {"level": level, "percentage": round(percentage, 2)}


        assessment_json = {
            "Personality_Type": personality,
            "vak": vak,
            "Emotional_Intelligance": ei,
            "Roles":rls,
            "Traits": traits,
            "Space": space_scores,
            "Aspect": aspect_scores,
            "Skill": skill_scores,
        }

    finally:
        connection.close()

    return assessment_json
