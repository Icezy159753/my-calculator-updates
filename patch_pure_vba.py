from pathlib import Path

p = Path("tmp_bs_xlsm_build/BrandSenseVBA.bas")
s = p.read_text(encoding="utf-8")
s = s.replace('AddGroup groupNames, groupTypes, groupIdx, groupCross, "Overall", 0, 0, Empty', 'AddGroup groupNames, groupTypes, groupIdx, groupCross, "Overall", 0, 0, ""')
s = s.replace('groupCross, "Index1=" & st.IndexCodes(i), 1, st.IndexCodes(i), Empty:', 'groupCross, "Index1=" & st.IndexCodes(i), 1, st.IndexCodes(i), "":')
s = s.replace('For EachHeader Array(', 'EachHeader Array(')
s = s.replace('''                If v <> MISS Then
                    If t2bChoice = "5+4" Then If v = 4 Or v = 5 Then countGood = countGood + 1 Else
                    If t2bChoice <> "5+4" Then If v = 1 Or v = 2 Then countGood = countGood + 1
                End If''', '''                If v <> MISS Then
                    If t2bChoice = "5+4" Then
                        If v = 4 Or v = 5 Then countGood = countGood + 1
                    Else
                        If v = 1 Or v = 2 Then countGood = countGood + 1
                    End If
                End If''')
p.write_text(s, encoding="utf-8")

p = Path("tmp_bs_xlsm_build/BSMath.bas")
s = p.read_text(encoding="utf-8")
s = s.replace('''        For a = 1 To 4
            For b = 1 To 4
                If b <> a Then For c = 1 To 4
                    If c <> a And c <> b Then For d = 1 To 4
                        If d <> a And d <> b And d <> c Then
                            score = Abs(rotated(1, a)) + Abs(rotated(2, b)) + Abs(rotated(3, c)) + Abs(rotated(4, d))
                            If score > bestScore Then
                                bestScore = score
                                mapFactorToVar(a) = 1: mapFactorToVar(b) = 2: mapFactorToVar(c) = 3: mapFactorToVar(d) = 4
                            End If
                        End If
                    Next d
                Next c
            Next b
        Next a''', '''        For a = 1 To 4
            For b = 1 To 4
                If b <> a Then
                    For c = 1 To 4
                        If c <> a And c <> b Then
                            For d = 1 To 4
                                If d <> a And d <> b And d <> c Then
                                    score = Abs(rotated(1, a)) + Abs(rotated(2, b)) + Abs(rotated(3, c)) + Abs(rotated(4, d))
                                    If score > bestScore Then
                                        bestScore = score
                                        mapFactorToVar(a) = 1: mapFactorToVar(b) = 2: mapFactorToVar(c) = 3: mapFactorToVar(d) = 4
                                    End If
                                End If
                            Next d
                        End If
                    Next c
                End If
            Next b
        Next a''')
p.write_text(s, encoding="utf-8")
