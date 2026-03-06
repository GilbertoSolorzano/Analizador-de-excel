def match_column_by_keywords(df, keywords):
        cols = list(df.columns)
        cols_lower = [c.strip().lower() for c in cols]
        for kw in keywords:
            for i, c in enumerate(cols_lower):
                if kw.lower() in c:
                    return cols[i]
        return None