"""Aderência analítica diária — Streamlit.

Executar: streamlit run app.py
Base incorporada: comparação dos planos recebidos em 16/10/2026.
Não contém resultados, usuários ou dados de produção do monitor de exemplo.
Regras pendentes continuam fora do indicador. Atualizar BASE_REGRAS_COMPACTADA para rever
os vínculos ou incorporar novos pontos, versões e frequências do plano.
"""
from __future__ import annotations

import base64
from collections import Counter, defaultdict
from datetime import date, datetime
from html import escape
from io import BytesIO
import json
import re
import unicodedata
import zlib

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel

BASE_REGRAS_COMPACTADA = "eJztnd2O20iW51+FW8AALkzZFr/0MUZfZKbTVYW1u7PtrEKjuwcNpsTM4owkakgpx13bC2w/wD7A3m3NXgx6AAMDFPbGl6U32SfZ4IckkieCiiAjGIc2gUZXVaaSYkT84uv/P3HiD//ti2DxxT8YX1yNRiPzi6+MLzZLb/2n7Gc3+Y8eonC3SX5w/X4beft/3/+f0FiExpW3XITJ71fe1o8Cb5l85NJ7IJ8IjfEvH41bP1qnH/D/ZRdsVv56m3ziTeivF55xkfxiGc6zP3vn7T+QHy4qf7j14/RvbsJl8p+xt9os/T9t/7zxk59+u4633nq7/7e176Vf837jz7d+8u7j5NNbb7uL08KFq40X7X969NPHrEPy3LQ85L0eveX+P9bzgHy5b8zDpb/1yD/W90G08sgbbfxlaOzi3f6nKAgNf2WYk+fm6Lk1ssYvjPuIlOzw16GRVF1orMgrBQvv2bEyWdWzPZQy/VS1qOEumvt/ug+WaVEvfv3s3Ztvn73+9tfPRhb5n/HUuCEl+reVv43COPnymyic+3EcJsWgtNOfXod/DpfP3i/j94XHR+G/xpTHm/9lGax/8EhZk8+mpfTX8z8nn3TeOz8kPwzvYj969Bd/+mf/z8kj/vAHFh1UIooQMNo+r4ik7ULSlF/84z/+96+MIq5Wl7heDrgOuF62wdWGuFo4RtfvVgSAhf+5IFssLgpspxpH2UJlsNF1ukRXaKQd0P200a0bcbnQdSG6No5R9+Lt50JtXlIUwM40jrVZPbBZHXfJqtAwO7D6KbJaN7ieY3UCWXVEWU1/aJjNRlU/e6ZRfcKhIS+j4L0YtCYXtFSuTiWhU3V8GQxcWSPA1csg6WCeGFqU5qMNgqyGOtRJqUUAZ9PuOKOPiANneDm7lMfZDHLmohrPhDUlhZhhEn0sE8VolldJLWTmqDvImg1mA2S6IOMayrggo7g4Y1wj2S7yf/TwcHZ6HRSoWTjGs2Ot1NNGMWFU0dZwSBto00ob38DGSRvFQ5k0o62hllcoAFUo0bAbYGsZuHYDtkzSzkpvrIbi2g2YFMNDFWdnR7WBM2ScsUa0BpxR3IkpqvGs+w1BP0INrDGK0YxrQ0CxFVRB1mwwGyDTBRnXUMYFGcUPmOEayXTsA2o4Q7YPgKFTWsYzzn0AxRVQRVvDIW2gTSttfAMbJ23QG8iUXBHa9n992KWv8tp79B78VVqd8dYPIo9BXlwC73pNvoIUBv55XORu840YczYXc90b9Xl1eYafl5t86fJUdL9Q9BLoeflRQA5jo6bvp4L2/RlsCqScASSrGKafb0FrwhQOpu6I8SsvuguDReRtw1gF7cXnG78yLt7eGoS4c12AfBZJHwD1g6I3wGAWjb2BtGm5O3xldPaltXONBd0bsxxq+448+CHabcICeoyO6BnxcQakgGM8efPHLy7++AWhm/zL5R+/+BL00gsjWJB/De6DuRcxFkL7n/L6oX0FqRljldWQcUG+6FJFR74hzydv6Zd7cV47kREn83+hS8/DKPLjTUj+yPDS7k16YhyQV155L4ybcL0NjXVCwansi/DFsbojg/y3F5POQsoTbvycmfou3KAh8PdpGwZ4MPt0lXPoG5m2Ss6vFOGd021cDWCfqfke8Axdfm6eoTNlluPUbklZvQS8czvSN6TBQuPV/qe7iDQQYPbWW5PXM252y1XeOABasv0JV3dekcgav4BgSGiUvxtdgXKImgY2aZr6lqdUatLyHlfLp4+P/McgJt0gRwCu2es3poxWZbUkbDxKm3FZBxa0qEwXEXBsTVcTb+eFXQ24wUVx97jxyLsWNKrMMSba6uQ2XcDxaG4amDNFw3KVQMcnvVnQuzIniLhjym0M5vgEiObMnZW/dPAmsKpSA1tWK/WgQf/KnLYB7SVpkWDp5b2qzyu4ckl6sIYzRWPaOJgrV4LEZRx0sswZKu50LeRqsEO5lDNF447UUMe1moOGljXCBZ2+9Vwdd1hXdA4O9PjWdDa0miwTFX2aVnU15KFc17mS13UNkONY2dnQUbHEHZXX4QJGiLDl4ysveO+lzbqLyNOiTCDNn3Fo1Hf7n5fBgqkQc8FGVYilKr2Hkpd4LL06LioFXPcqKdCTsMQ9CRWkiK/9UVCCdHsg4ERXCYEqv+WgIKTdlKWJD4yzGuXQOjcdUJK3XNV0vPTnZOpLuDDevf12Nswvqvlo7hHaUES3xhr5eOmv47rsaZgBKb87LkKg/MRNCJS7LWG5+zaMtukc8SpYkgIDVrIfx8aTr0m1L79kbIrywCUj/VBpa07Vg46RVNJ2Rsdi3Kfvazx5OL4ufjGIcght+jvRvRGlISltV9Nk51KS2lD0toRFb8W4UTOP6aPtbHowHbBB+Ud8I94atjMJw2wodFvCQrdi1tjJRPUBx5fxUwd1FPmn+yGOJweoDeVuuyx386y7QIMF60XwGCx2+Utyrse+Pf5VEiCerILyB56fY7m1b/VrsnOV0YMZWuxsZQUpB2rYdkMNO+Xdo6zpcwH0eHqp+MEjTre7u93yEPtI+6QsCamBqn1feI0e6EJiJxRq9exiA9Q149nW4/KQHahw2xZqGIXHN3UsohycRALru0KRx1h2oIJu27hJbGA0K4QRqcEsEhjfGY98brMDJXu7YWB+R0h27j4zccSoz9sUdU142yEZRQ4X2oHegN0wXL8jDGW4BzJhxGsJ2DLEPslIHmrLOE8mdCVscVciewPCmO8taqKsz7mcpCzkpZNjZPJGRPW7YXbh8Q+nFO2QxW4VHGhW2JOW4NSEcn0O6PQsAswWkACr8EDrwW4Yb3+VVFjWBuyB58gI+Ju62TE+Co/zw9/PswYr//0ZkUW+enxVfYu+nbO0ofzWdM5ktX99kzds6UNtMsVmB/ocdsOAfkRgd2X6CnCNU7iBia16gfU5p9iBForT8MQAJqpZ4o9esJGKQA5FlOwF28f6ZOLtQjvHaWjnIML7Yk7Wpz/y4813gIadCy57KFnzBoX9psIccQJ9qlAVuPqU6L0ieLrVqUpr5QUXmlNOQ3MKUddKXPy5b6jpYcZfjF/76x92K0LLQ5DuIxdJRh3DX8ce6TikX2zSPWhx+9m6i8Ai4eopYod6qxRCY8ppaEwhopAqbrRau/RycMcojTjyNOFOR3UOK8OFjprT0FFD1Jdud9HdsFyiVQaufiX1qHOnXatYqfUdDHqFjvg5ovrmbn96RLzDaNLWew68wCnrKkfQ2XMaJu1CNFCTrhMu9x/zcfZzW/ww48Qr1YKL4Z6K/eU6rR+zoRnqNExUhqiviRy+6edyiFXZ6A/9OKJXaODpVll11ncnaA87/beH8wTwn3J/4lNj83rA1aGkpv7tuEOdv1jAhb60I8GXrgnY+eRDLnqXGNGlhHFLWIeV60FP2IULDWpXgkGtHW99gRd9y7/oygsM7xjuc8EXY+hOuxLcaf1s6wy/6GGaR1eNp9AF4edDMMbQJ3Yl+MTaIZfv0Sm1h3sWgew2D18fQ0vYlWAJawdOlZGlCzu80rwLpXmBkIQxtFFdCTaqdv7U6fMdEdhLid0VkNirIEK70W14NBEViJ++eM05gmKUr11l8nUHy1kOAXsMvVdXgveKoFN9VhJ2ba9CKGK7ykTsTrrVeRl7DH1WV9hnzS6vzBKNGW9DQmjwCI/TX+e3+VI+d+xNFytCaJqV2bgiS4/9h/Wu1F3OXAQqMVyBcmm18tvZyad2y23wT+Q7T3dUk6/PkoK1uphaR9+Rkf+JiVYNTSyIwCXV1b4ATVJX2CTl6wt5Br6hK/B0hcN419ueMJaRfooJFpulxh0Beptu2dsUuEy6cJGxZ2z9Zdq4eZrDTZRsNUHvYMdPlmbC4+OOXBQeKLOboL82mq+O+9RjKNI8Z5TmGDqXY/HspDjgrckZhCQmuDmAePMPjQXyP1fgm0BrcVy2Fo+Bp+TNL0i9fswpgfumm2j/t6fXj94mpGanKv+6hjTyPcsg+avjle5x+mzj9Md+XKSusxCQtAh+sYSiQR8Oab4KHF4y183T5TO8wp4PCud0hb3A7Ud/YLYtqz0pTdis5c6Gc0ygGzi2+sZlV7EbZ7A8H62hnErKzoptXGuj8lwcxgQ6hmO7d1B2F3RxjkueMAvlaFJuQ0fJ5vkIigk0FMdO3/BEf6L+DNRcx+fVQy2WVqIKEjQEx27fQMrTGSx6kKThDFLUkuiEq1Umhgk0xsbjvsElEtSFDaezYVzqARK7Kb0KEHSAxpNOADLWv3wcmXUYZcovx9605HEpWnAZpztZUG5VKbeWq9urntqO2WRcd21MoOfSJX3WQF9v6bMk0AeNji7pGw/09Za+sQT6KE5Fh/RNB/p6S9+0PX1TilUxFabvd14UbtJM8WBxT4+my/6ghN7pOsLLcHXHAx1jazBuDlxNOfDhBuPknfeOKG41JaY0FqWNzpoOU4rpMMNGWFemAi9gCNwFkxL7Lj6ctebrnH0whfbBZIQOr+7sAW7CcPgElMzuOiA77wNMoQ8wEff0FXPWZepSVaF/vADrV9wsoYA+ZfByJBedQu9hIm78K4b3+yCeh7HgXeFtD0PoprhaaJ04Q3vLfG8smIcclAFdqZN6sqHxMRGPHlBMdtdXBuiGGoejZkFHzep+f8SZ0H8K/ZeJeJiBYo575P7yoorNBragi8eCtkoQ9FAm4vEFqkfCDk8Pax8Fzx0QVg+TaFpvdaPg+bO9U+jCTMRDGJTz2+FBXf0AnzuLq55gsWQgSgk+f4x2Cq2cibiVw4vwy98/nY5GoxK6Cvwb+XK68eTk5MB8IBi0dQtqn3K19ULbtfNvZtC/UYjc7586A3H9JO7UdC2Bg3bOpGwYChysy8/3Pq2cdTeOnm9kvLs2TDN/c84TdjdBPA/W3unscBHHMydBuY1s9KdBBeu26bFQ9T1DwCKoskrxhhqfYC7V5/Ew+oCqbFTPVi1eUgWOCVZJhe7SVNzFfOPFMWkxuNkJf/Srodr5zwzyrkaShObRjwwvIk9ac2UCT+Aj5BEAAXxW8+k7K8Afv7j44xegpTlmaZdyStgzXnmg4viblzzS+vJwCvSFUf025vlQq4lyXmi/QpOdaym+aRs6QFNx+1IFYPTQi+75Oh9yoRivGRMveNCzQ7zyeqmnC7owU3F/UQldrMgLDYDxRFwoZswcMSGjHNnskrJj5dSDBm2Sqbjdl73iZf/nyUuU86RpsSGjeHLNIbtUMlNCH2Uq7sSpQEzHTEkjTPtMadpswCj+WXeAcc2V0OeYivt0SvjSM1dSEcMwVzpsyqBa1yVlnHMldCOm4obaG+9hRQXt4JrkHzBKyIkH86tiLHk5tBMle8lvSlvzH5uP2mI1Ef11cJkj6DtMxX0H6XB1O0Wy2dI+RVrshT4lbLUrtDjmR3MEDYap+Ikk+WR1PTnWwIVgcrRMNl8wjrQzvrhmRnMEfYGp+JGkN/7SuDDeBvPQePX7p1ThP3tv+gfxzZXkNVPlwogqr4ps2rTY+0tazGcj9qhNVtecjSdSqPzPGij/6lnsdGrlQ1H/LMveiFIil/SRyDXvQodg1sAh6ADEjmdiThYxTMrsHSslvkQjjpzTNHQVZg1chfStbsK7yD+DZLK+oH9YzVzd4pTnCcpN9V2xTdYum0mKXy6eJJLRZOeatPF8DR2IWQMHoiMohSftDpjUP2uzjXxLSk5deUhyTdzQsJg1MCy6IrLB7N0FlBimb7YcaEvJpyuRS84ZHHodswZeR+G1XgbLHXkbmNuIWojipxHvtzM0F4WXRTaJ22w10Zal9jCa7Wy7Np7IoT0ya2CPdMamnv33GTS1z+U2W4e04ek37WTyzOcmtFZmDayV7sDUtR8/xyaCKd1mS5W2VIFIEp5807oJLZpZA4umM0Kpp45r6Oz8tDo302fPFKvmmS142lBhqkkfUiUKOjKzpo5MF0RdBesfPdahjN5RVSiNPrLY2qUNTwAJkAX8lTx2QZisS+O/koo3zNHzLNNnmarkEMn+w/3uwat87Mwd0Q9+5C1JTcbpn/2///G/eC7hUrlXviS85CUw/DypKbbtCFtTtKVpitXmZrewSMNybU1MYMXkURG9YLZ7zfEssvq3KWzJ0ZYmOaohlmvLApyaPMaiH8DqkCTPM4th+8JWJB1piqQibDm3MsDOySMymuZPyIN3lSe/VsNmoSDzYkGQLQActh7pUO4MFueS1aBiObHryQO2TR6CgYa8LidyLvC0T+MOW210KNn/u+aOa6oGlkweZ4EHu26nYz7yEEzGDltLdKCW2D18nBMusF3ykAo0/KFXC7mI1a0UOmyl0GmjFFrAHckDINDw0wttkIshDLqgw9YFnTa6oAU8jDxaockxEx61uSfH5DCHIThs3c+RloIFNqjUg3MWcDrySAQk3Gk6QYc4xMBha3eO3FOajajjWfRb0AQxm6SXUgadtsN1uMMHHLb+5krLpNECPb4lvwXtDFPczjj8nI3frbdOjsyTZ7z1Y/KJHc64v9NrRvlrYpTYXLbE5krKsUFtUXojNp5toTFhihsTKtHrcr7lIU/7dOuyNTZXUkRfO/C4JlxoK5jitoJS7rqdcrnQQzDjumyRzZUUsNeSPs45F7oLZtld4El9e3zV77//5sbYROFid0zmypnh9jby1vEmjLZpHtfIXyz9qMjhVZbzhvNGhTKBZxLbPvpRnFiFj8FD8qnj7QiSk9jWVVJZTsmLqg9u2sRt8SSoNS1oG5guRp5u/VXSXNtdRBvcPkmuqkXWxxdtfubkC9oC5hgjX0I3BY/7jJX+laDADdcVnGzoEpgTjDh9txK88LLXSBVLqw8rKBVzYwVNA1P88gsJWF1s9j/Hxo5UZ95w7aASui9ggIoGlUAUZxUq6AiY4rdUdLFUZxmbnyJVGLzPsUCMZRUqqPhbI4xQXTx78+wzQeprUhW7ZZiAEnjGxRvNcMGD29xwQUHfMjHCdfXs+88TrqvvNcMF3SluuKBkb1kY4Xq1y/Mxf3Z03R9Lrg8wgTDvKmBQm7dsjIC99N9vyUcEdKxeM1Yqrj6wBOK4q2BBzd1CqblfkB3T5zJyncqqDymB6OwqUlB2t1DK7t+u43C5/5h7gU3ux+4ZV5UC64MLqlnZDdkhF19QdrfKsvvNsQq4nOptQEgkVQfQKgP09ggQ8KrD/JfG+pePZmkx38IpFPKqTwwg9/PGNeFhtHyvDCvmDzWNTG1YVlsym/DqcLFJ3lZVt9qBcr01QYqhDINRCY5obMBJTegYDNnXiGWlxth4QtnfmiLFs4U/qYRK7S7ipCaeTMAJUk9jXlNsCqFPYM2QUtjS1lRCIgqfaMKOL6MIshppLNQWm0hoMthtTIb+uKGK9hx9AJh9rrOF6OtAR8EuOwp4xrZ2NqianQgCs3LCPqgpItZ2sBk5VhZ7YIMehG0hxbGNgaoERlw254R9ZlRE6lWP5aEd2VBC38K2kULZxnjtAErd9uiEfaRURCzuYKzM25ENJfQ8bAcplK0M2w6oRGCrTtiK4hgqihrBPLYlm0xondguUjLFnV6+hCGt4MThx07ZYuIYiol11od6KIs1xuYSWi42VstF0CjuAEoEdu6ULShOYMCmXiKP1cXE0YXWi43VemlvMithEosVPGVrixOoLeoFs1xnbDqh82K3OXDRVGHMy/H2dlra5VAdae40hXiURe1O9pStKk6gqmgbj/6Pfmx4IW+2Lxd6J3abMxYyw2gMy3jyzguybvIuXPrvv2wFGUdYzWcKGVuhnECFUiCKxoU+iKPDB2lAV300Qw+HMjRREFO28DihCI/vpz/EXKxBp8Rpc/aiQ9bYRlwPOUNhxk3ZOuKEoiNyMwbtD0c8a9L+rw87z3izW26Df/JIA115m21W8cbhzwFvV+Fyt/aMSy8KV/u/bQ+fOVJ3E8TzYF18avFRx5nGi+7CYEFGgVC6T3dNXvfRW+7/Yz1PhOTCNxm/Mi7e3p46nWds/GVo7OJdOpEYpFuak+fm6Lk1ssYvsuUzqb456Sdbj4CZtlX+3EN0q7HyCM+LMKZncc2qOHnIqW7Jn66S2nmaVs99oXrKs3G1jvRBzNYcJ3TNkQ4xO+MOJ4kM+OqZI01ev4OBro0jnglKX1/afCO7B8kmOX9DbfzO2ArlhH4qHBO/We2x8YX+jiN+S0VWAPJ+z0pvl77IM+Pi+aUwvO9Kd78Us1gfS2dckEH18jOeF+ZZHd4ldfjMeEKAT+v6S8RzwYytrU7ph+Gb9aUzKDLoE4Lu7MQA/SlH/CIOtIusmiU9X6wnwg7V84XWjK0PTymxp79zUE1UeX8qcQX6FPTWHPHLRZAttlr3pE9uwcUWi6dQLMbGcVZ7tck7x9CTc5reAP82mJPv//1Tc5Rde0tJ2E7/IE8S2e9JScO4garD58p9uyVDeWW89g0zUf+fLwKvoG/FdVPAuduRo0rJS7BXi9g19QeyLaHLTs5dF19t7DoQKO1fqZR6mKGF54hbeMmr+evHwEsmbC/yjHs/Sig+KXm8umX5DskVeW76PNJa2zQHb8v7l7t2X85XjGhi+TGd4+JjE4JDLoLHhbG5PGqnX/PkpAIBvDMDhsveG0N7zxG39zQxJpwBHjli50/8qSJsVksY1GJECIMWnytu8ekirEGud+yQ8eSKV8WZOaoFzWRFgfGRBg0+V9zg00TaxZw858dPirRCkTSQZtaTxgrr4iMN2nyueKY1XaS9/bQoe6uLMKueMLihFiEMml+ueKo1bYTdfmKI3epizK5njBV8xccYdKjcikNVrZZqJb4K1t4ScJRfybL+5WN2L/sRo7f+Ir2phXwu8kp+gIIdpNjFdgkR8TZMmKq+EJZtoOnUwwBjo+pg+ANs3VKLVhuRq+24rhYbQ/9GOnfjgbv+czeWzB30OFxXMnctxruWaoYa7PRJEqZbTx3rCFuHox3PfXYT6EhIh67FYDdAhwM6/qGOCzroHLhjydC1Gelaq2qKuNMqjdWr/KaYzK9mvOO7QXECPQXp9LUZ8gb6ENEnMPBx0kfxGyaS6Wsx9rXXeVn0GX8xfu2vf9itjMh/CFIhZEG+NHmv2AtCYx0am1REKeonbIefi1m9Em+9a2W2sq0mFDNBNkYtBjFsGA3oSUSP4i5MEY1gCtTd89nqW8KiTcG1xCxNJeu2QgAoe9qkGA6yoWsz3g3QoYCOf7nGBR3FgRCPPiLtuEgmGtIct2EUpa8W+fF99CypzcaWV+lRgZf8RWiAijobgHwsfJnPrk/nClaSWNiyMo7rLVaLchHfe760uRPoQYzLUUn1w11IBgAI04r8OEqPdZMm2u4/rHdhianX6ZDhGS9/B2v8jP0gP/v3sQxYfAar3mewoM9Qk+UbjmiH4rJaqa5xDrXCPJM0gd7C2ETBk8gxVI4rsq+zFT1pgTAd/dOI8XCXjEHZEj85CuHHq+RcUFa+4nj1SMow3y0ZEysVSH2DT70WbDW/Y30KPYGxhQKWbFNlxLvl/f7j4dyFguvV/dIRNfCtxnPjtL3j5IT+6hqwGddj0/wO9SlU9cc2Dmyoq/OupiyN6+x6GcCCMkBnM9a5Q7RTKNKPRYIw1MH0bv/zMjlPyk9UG3WeSlTpFbrHyq6Pkrbb7+7akHWoHKOeLyjDj8XvX2Rgxty0vQyjtRezd2n7v87TV/9N9EA2ZcG8OWU4Ll2k48sopQaQYXyiaaRtRBaLEW8KwimU4sdlP/FNMI/CuyBchg/ZI4t19D15ldC4jHYUkrJ3MUsM3QRxSJYRpOKD5B0pm+Omh45bjFNZIe4OhcCxQrbrN2yU1q8fp0rNmA4vlbYrNNf5VuI4LjyFSrtcsqyBrD6RZckjC8rpcsmyB7L6RJYtjyyomcslyxnI6hNZjjyyKJq4VLLcgaw+keXKI4uijkslazyQ1SeyxtLImlGk9IlMspruDa/2f1vull5sfB88Uu+81BK7UAdhWlHxPNyQlvu7NL/VoQiPxyJoIBTaLF1uImEz1uNIkeil4th0QzngiBtHzp2nKI4UkV8qjk13oQOOuHHk3K6K4kjxBKTi2HTrOuCIG0fOPa4ojhRfQSqOTfe7A464ceTcGIviSDEjpOLYdJM84IgbR87dtCiOFAdjKhPHpjtr0Lh3UZjkNidNK5pPHQmdoAAa2IT3xQicWJpRPAmprDTd9g6s4GOF4jJIZaXpnnRgBR8rFN9AKitNN4wDK9hYsUYUJ0AqK013cwMr+FihyPRSWWm61RpYwccKRUOfyWRFzj4oree0mp8fdrvpZrd/0NSWRAM9Yjm/qvRQJG+p9MjZGQ304KSHolBLpUfOXmmgByc9FEFZKj1ydk8DPTjpoei/UumRs58a6MFJD0URlkqPnB3WQA9OeqBGPBnJpEfGnuvWe5/eEbgJN8W7MZY+Kdgu6h1CPMXRwFGb2+2sEdSP5XIkY/c1cISeIxNqy3I5krEPGzjCzxHUneVyJGNHNnCEnyOoScvlSMbebOAIP0dQnZbLkYxd2sARfo6gTj0xZXLUdL926cXbcO1v/V6Yp5W37Z4CR+xWVNmH7crlr4tOtUyobcslrunObiAOBXGc5+lEiIN6uFzimu4BB+JQEMd5ZE6EOKihyyWu6W5xIA4FcZyn4kSIo+juUolruq8ciENBHOfBNxHiKAq9VOKa7kAH4lAQx3m2TYA4i6LlWzKJa7pXfUX+5qAIaGDuhaEwtX8du6DcGui1hOitEkVR9aUS1XQvOhDVV6Io+r5UopruNTUT9RmQIHu/Wa2B+rmR4gdI5a7pjnPgDg13nLtOMe4o/oFU7pruOwfu0HDHufcU447iIkjlrunuc+AODXecO1Ax7ihegsBNTnXc9esOwlpNQ99NhE59WlwHpsUVudep0nwq7iO0LIp3IHC9k3LCXifhDrShrVvGCq+hgbL6CwaddvddSqDsVDtsziiOgYuIs2syCIfLppwJ3WI4D1dG9nWGb+zWwcJb+MJA3vqkuN5yHiaXbM21kVl/h6EjcPWlGjJPzcomk+IsjEXJvCK/j4L73UN+NxmqGzc5rpCWdONmVhnzSmWUwMVy/aYzqUcX5oLgvDLasinGwQQfUJ3PrLVY6J1gp/UswDNq4sNYtUEVzbM2xWOY4oMP73RbSymSWbf+NllH4OZgpbhyTL42xcCY4cP14m1HFxfzTKCHd+kePLf+vlkXevt6wDtUERs76F9MR/iw+5oMMn4kMk7aStErvY8G/Mx6/KCwx1yyKcWvWE1sBKGVMTXxISim8KnFT6PQ51r16EGhTw96Z/U+GzoZUwsfdtT71rRAp+3KNbde9XOh6qcHuayC2MBBC2MqbGGoB47rwveSV/NVp3ZaLaNo7nF34dZD/B53y4aexFTYk/g+JH+ftQDFby389ojJq8j3f/Sj5JDlVbj0t55RHHDkyCbywXjMSmKYgAqtkoo7k7AXKLQSaLb61uJRSqAjMRV2JFRRplkfeWEoYhKHgDIeaYWTRxeBpsRU2JTggrMSNNzvQdACwGkdBMcyBBEWZ1lhWw+DDnQrpsJuhTrSPo2BEHKJZCCEkkm3gHIMhQ50NKYCjkb+y9C4jbwtbftwtbvzSmBehMYy3zEkvzPWv3y0uAa/mmg7IQtWCDRa+RCNgO2EEVrhDu3FaqZTeWtD6xzoPEwFnAf5WDFHOllYKRztzkKIZLhrp5k0ofHUqvU0QkNiJmBIyKeRKsDhHOC0KXRjGN6kGqesrPUoQWNhJmAsyEeJJ3JJzRjHjl5SN9JhiW8aQymON4DJgR7BDMXJ5+GC2ZZsFPwDG86FXabwErtk1nKgiyAXyaZHpwck8SPJmeNLFEnoUshFsunZ6wFJ/EhyHsoWRRJaGnKRbHose0ASP5Kc57VFkYRGhlwkm57YHpDEjyTnUW5BJF3oeMhFsulh7gFJ/EhynvIWRRJ6HDNbJpJy7o4d7hmWxyfUDgXyPrnQvJDLi5zbYgdesPBCsRek8iLnftiBFyy8UDwEqbzIuRF24AULLxTHQCovcu6AHXjBwgtFzpfKi5xbXwdesPBC0dodmbzI2R8NtwSrJAhm5BAhiCKNSyVIzo5pIAgvQRQlWypBcvZQA0FoCRpThGepBMnZVQ0E4SWIohNLJUjOPmsgCC9BFOVYKkFydl4DQXgJomjJrkyCZOzFhruAu2IJnvMWYYmiM0tlScaubGCpHyxRNGipLMnYnw0s9YMlij4tlSUZO7WBpX6wRNGupbIkY882sNQPligqtlSWZOzeBpb6wRJFzx7LZKnpPm64UVeYBAcmNOryMJ/ArboTigYulbqmO76BOjTUcZ7XE6GOoptLpa7p3nCgDg11nEfyRKijaO1SqWu6ixyoQ0Md56k7Eeoo+rxU6pruNwfq0FDHebBOhDqKki+VuqY704E6NNRxnp0ToY6i+U9kUtd0D6v5Xt4XxvU69oLQIO0XrtPk6Ia/MsIdAcXYhEn+IvKTlR+vyP/nFfIiyZ5+eO7+w5q8vzi/Ou/3LRIME68KKCMTivovlaqme9SBqj5TRfEBpFLVdA863CHeBQ2y96FC94hPKL6BVPaa7kQH9lCxx7kbFWOP4jNIZa/pfnRgDxV7nHtSIfamFLdBKntNd6UDe6jY49yZirFH8RyEb+L9zG4xAtzpSV5fvGZNxt3PSu8xmlJcBjWcfUpXxXxmpEm6LGZK8RaEr2se7swS4VLzzQlFQGXcK6701qwpxYRQg+fne5fR5wqopNuMpsCxsEZaL/rAeZsRwukZBlsiucloCuwKa6T1wo9P4SYjxMMc5fZKJLcYTYHFYY0ELhf/HG4xAlx1f3tRESWopyC4wWgK3AprJHBl+HCDUTMStd9cVART4Ka2Kj3Ab7BG5QwHL0mBgyV7D0B+4wFovvYjb1mi5p23/0AKvEhmoOVunR4SeSz8+bFe3769bbroTz5E6venR39Z4sX4i/Frf/3DbmVE/kMQJwb8gnydb/iZQ78+GPJBwlpwT1pzET5jru6TVzYeDiUsU3F4ezYIk2cjC4JwE4WL3fG4TDaDJDx4XDxMTjdZCazqKyjMgPxvjVydKLTbAbJguC4NGeSrK2sV41D9x1gM8qIbfxkau3iXqOdpGIc5eW6OnpMJe/zCSGv78Lx8MDFWZMAmi6tGFImtnhQDJXQ1fZWQIhVHEM61P8fGcAbsAms05mb11dJ/WLWA9b749z0ZuI5lRjtymQL3zFdxAKq+NZroxeHTHbzOkYRq9DKFLpWvUlIio3b8KjHAM4ABb8AaTbmJzdv4m2DBWLRDcN/65C3JK8fZyxO6PIM8hVRL8voeIeVHf+0d17rHs71eHHvGdbx5ZlijXz5eyQaaSli1eMbXVNCyd/Nj8pX7D8nYZ3jJS/7nlWbkhK6JrzYzrWmPxRdsRFrjsYkEdoA1miEk8pf//PbXN28wcYhrwIN3yGui79RQbOagvm+OUDL39euBOBZxUH3QRlzWTGzeoPhvmgh5uwrXi902eEzVeUzgwRfTh51QBIhK7EClsPmDkr9pIeQvEzcxgVd4I33ECTnmKok71QYbNWgLmDZC1KimkzbMzvpLqhFrpbbJRCyrCTZe0Dcwy77BO39lPES7TXiSIliY/XBsFT89KJfurTf+2g8iqMlcFMSuqETgNfWPBXayNa4UR3TRTarErRMIinJc8Uwf+W8vVfF8I9yQhjm2HIvXczXTh/2w1VjXs0fQkjBd1JSdW7n1HDFE6z9LQNmrYgXdA3OMGivmgqznPGFY1lkCel0VJOg7mBPUIInG+PQBIt2LNktAcqsCBG0Ac4oaoKsw+vQIOhRKH0LwkgluhKBub85QI3SRrFBJg98H8f7DXDCaug840QqoDy0BnayKFpTnrVFTtC7WpAqhkMCG6CarXcM08uMEsj3JzuA5Fb0X+7XGIWT2COrrlqkfmHaWoW5MUJk9loBAVYUDit+WhQGO5t7egEYBDbuFzgPFasvWj4YUG043I4gkG7uFZAP1Zqux3iyPkPNL2V4wgmrBajfXY0yoF1vienEW9We8tJSuPRDtcUoFRjzBuM3BgIqv1Vjxvd5F4cY7OYiciFxsvMhf/pBUip+GcRetu54tVatVgBca2jH333FCA9Vdq7G6qxKaIA43ydNl50vonJxqQfRRQ9nm8FIDJV2rsaSrkpr1008Dmko5tDFDOWnOzQzUcK3GGq7ikeZuJz8xi46BplgOfcxQrCNeZqA4azcWZ9WOM58EMuVi6COGEm3NSwxUZ+3G6qziUcZb7T8cF4+9ZaZaEH3UULZOvNRA2dZuLNuqpOZdSL5w/3NyxpF8av8T2YCEQXzMUUq90qRHLPEVTx9hFAeblzCo/tqN1V+VhL3xP4F5rFgIfbRQTGleWqASbDdWglXS0u4QBQ5YMMTnjZsrOBbUg+3G8cNKWTkI8PsPcXKM/zGId17PB5qaMuljqbmuY0EJ2W4sIZNSxP7Dbr0Qiaq6PZ7W8KtPSSopPw2Sp3eXc+yBIxmXCq7K1dOHoJkJJc3g77hScNkWlJntxjJzd2A1Mbbw4YTKq5hANYgbIqg6241V5y4hogbfDAg1RoiSCpAXIShC2wL5wPd/fdh5xpvdchv8k0fq/8rb5J+stH2Jp9soTNomyxHtLcOKqx7E82BdfCodoysvugvTCN1QYJcvlGkyTt7v9C2EpYjsuTfhOjk8aVy8vU2SccVBvPVXHhWkrH7Ip+enikkuAEyK9jQtW7Y4PF7QUo7iqZZQX4LBFjMdlKwdgXTeOgkTOabMwdVtsEnbYZ6lRV8Ej3704Kex5X6eg8tXSJPuEzNTSiZdFkOUTLrcIFDb/lyTn8+5a1tQSXcEkojrBFk0L530VOOa0T+uOPTBD+Pb8MDPl7vctqAp4JRNge/IQ9MWj+n0vwqW2TmaNBEBeYlVQNr9rU8YDI61vE0nWdgTjkvN/U/Js+7zZ1U6RAqJtyJQR94DWf0+uSBrwYfQeO3dPTPSt4i/PDvGl6qhrmsIpaOo1AH53/XthfgwPaUklfbzohnfrhc7UvTAC2JSs997mzDi55U82TosKSGu7Kzn5WZv1NR1rcvfqDyjOPQdHLvvEN/uoruas+K6UC69lmagoZ7cB6CLNViPNTRIHKfvWPPk1dDFtsgJi04AF7p8Bw3goBprKbehteO4vaecrNJY56Fb0V3OG51+jfGr/J+vg8fI584XneYb/+qwevfr8keHMT2BNGevOtaF5t5EyUvei+50atzznQl6W864753pXRp4xbp9htGdhC6MXoekCPMl+XTkPb/3gvdpFzjl5j+/iz3XCYpl0N0N4La17jbpKmHQ5HImfSfsYnnvPV1799sfmBe4cWHWwfUPfLiBAulmDu49RZiDnpgzFWIupWydRDcEP3oUN+yK/Gly71THwoNptV3SLtLuQwarpR8QflYErtB4Mk+LExsXX11+dfUlUjmCkq7Lem+JTMeUVi02pFSNwYammhoGLwcGPw0GL+UzSLHdlDB4NTD4aTB4JZ9BimM2U8Fg27m4hdCkCUd08hMlH6KuGVpQU6K4WkogbTtZD5CihlRwCheElOJaKYG07Ww+QIoaUsE5XhBS6EG5IxWQtp3uG4mETe5qzL/I+NXx3y5TwbBz3b1Rz0CkQ1pC100pXV2UWvJsn3CgY6WmT7RdXQx9YugTjRczgn0CGk9q+kTbxczQJ4Y+0XjtJNgnoFXmmkJ9Ij2PmbzbxTI4BY/mcaCLJE9hbDy5Ii8Trr0vQY85mWVFc630+S41vrY7gYdDZXjVypifKmNeKBxGrQ+m76jzweg4C1PBAYJUjdCBfp0rFlSMGfw222G9fQDfthhGcPatO4htpx3oIrpikcqYe0bvwijadSdsQRaUlP4CQRYONBddsWhjzGi+3EX+j56xDbfUbBx4x2zw4poho1wO0LchO6/S20OV1o/Y0O90xcKTMXeLV34kGqfcy6H6VE7d3QcemBUZo6Gv6YqF92KG8bC7/i6JbTztsXszVB/ef5e8/93x9XUTJxbui3HAZoBRP25Dd9UVi1M+J0i+TuW0N6mctv7lo1mi//rZ7bMLo/CRNtJJ64NPNAkQqUJiU1LniB3SoLQbramoLcQjblAcUaVgWQNYfQHLagOWS7EV5UZ3Nh+xdJ7a5GIMnbRlU/I3dTeOialSLsW8U0qewJA2kIeRPPZAJ0gexSKTGwPXfMxrvR3phr/Sa+LZdUCjq8Phr8mGwaWYVkpRFBgEBxSRo8geDxuhCF2isVAwTUzZqpMfFvbqpEgeKdAiuW8A6jyn3+ZBKzfkY3Get+qUKHR3t1uelIDY239YpLEpi+Oft9mctFZziqW4wborgbctCKszDVq7toE525VnSwM9pbFQBAxKkHUGANQzjW8tKmZJYsJbcCELbaKxUMgLStI1uqb1oGNzRR0JrqgmzoXcUBcaUGOh+BWUmOs3nuppLxlLBpqFtSPmZWKCvtGqHDpKY6EAGZTs9ycCoL6PoHH4nVbplVzoLo2Fwk0kQvbd7XWJrNywvTgZtvknMOzw8ldBuMVzoPHU8TCZ1w2z/Th2cmNoTo2FIk90colkw0ZDFN2OzYHuATJaxTZmY+htjYXser3gRqSZl8Zv3pnWzFJBbiWfaOn7ksSi+Q86P7R2ttuAitHdbySc7FHdb9I6O99joCc3FnKDdfYYCblCmx7vLOwl0PUXRMc5He2K97mO0mRzOIbu4VjIPdTZaXq5B6RhjmcT2Crf6Rj6fxNx/6+MztyLHpIJnjTnh/WOEtt9+LhxdXll2GcAIuMa/Ly2bWDSUU4B2klhn51Kajwhr24apE2wHgp2ZYnH51scNhpn2/JsEaHZNxE3+7RRq9HraMwvNhvElaUIKyRZyO8YQ1tvIm7raWP6Ykn+LsgubDNIqwf9OJbbuD8wyqu7V7RShMfQcpuIW244EPwhWET7n98HCyUbNKQcVgqtG0YxNboKI/TAJuIemDYYdQrBjWFCJxG7siRihbO8oEgMXbeJuOumDWuNotfFcqsjpVnTvoRIB3NlCcYKe1Gxlc92ogm0CCfiFqG+TrRb3gfbfshgzfkvFFI3/2I6cBU2aOtNxG09bbC9CuN7Tw1sKrApvq5ubMSUVS3DZqG66kdM6LRNxJ02bRA3u6FYH8WYri4e90Ap476oeALNr4m4+dWU429uLrOLcjk5Pn5eo19xBBjym74eVqtiLCGZjCi2x/bibFYOq2ICDbZpdwZba2D1yhji7KJTMMayQjIVYiymYEyg+TbtznxrTbRGBUNXVvYG/QiRejHWoAGK9iCxbO0TaPVNu7P6WvcgXTtKcYoRbSbHGjQ4UYq5N5PQJ5x25xO257c3MWiNmMcSmDZup7pB+2/anf2X1K05EoQs+wO8Wz1zhHavJyE5f5PBMWsx3pbl2e1Bd2/anbsngVrU+z0awPg2fLLyUChlWWzLN4V+27Q7v00C1sOmj6szIdr1TTTJf2LdSGzfN4VO4rQ7J1FCN0K686OhjGjrN9EkAYqhzLv5m0IncdqdkygD4k9n+0cFH8v+byKm21Uxg0bfVCxHJpkuSQG3wXy3JPVsPLm+vXj9NUQr/XEJqEpak3fHtCbZTQNtNnntk2OWC0U60tKLcpSMY1EQbvEmUA0QTI3Jbs9juTlajmMTN6VYdmLonRrFE+Du8LpkORGu7ryBMz2cMRrvWGhKM/FABV2zmditxV2MZ+/2Py+TnQV+2kpvqps5GCimZWw71IlxnkboQAnSKH2IG9DTg16D4U6AM+gUzXhjBV7tlkF6VVMDuq7JsjtZNaeXSRUp+27V7e0Oh0LQSSq+jW6QoATPDRKlpSj4gDYplL4eImgFzcp2+Tvy5Q/RbhOedLaQMYke0xS+9OfL8JijsETWRWGHFpUIS/OuF3m6Is8IjWXwGLVi6sZfJ9/ol4C6STeL62TAL+4YXxwLGRlJr003mr4Rbki9ptND3WjHKH5ZUq+USTeZUFBnklklB9oxM0Gj+sltGEX57XLxPfllAklIua+Q9jHjgjTIxypE9I++2v90F+Xiq1oP0fhLoiNUCVokB9WTF7v92kkRiwl0CSCx4W38NamKVLyISyyePTq1TR6ZUhElhX1mJM9beet5RsXxG42/N9758zTbCvmsV6g2hHuMqZQsGjVknYfpPEMce5IZtHW67R5vjiB8TxYhWb4hnq3LAZzK84auM3QdMdDE+eLpVtDmmQmGrmiddZL539+K7sQ++Q5Uqhbd3UhKZjvFM1Cxwuo7DLSUuu0wCuehoTMNnUn1nMTf0aCpNhMMT9M7M6mOTutzj8IWCDeVcnxa9RwlFAc3g85gt/1H5UQ19K2hb3U3ZYn1O4p5Khh/qnXeOoQJenlym6FrgQunjlWju2dJOYqgeNYSSpY0o3i9nXYehZPW0LGGjtXVlCXW6SjGt2Cot9YZS2XCnR53Kkx5faZSDhopnqm4E/vMKCZ/px1G4Sw1dKahM6menfg7GiUmQvD8htaZ6TfRNrzPzqoYBJdtlMf6DB3rT7QqenKooy8xdLOZlFwuiucsFmF13coZUWIpOu1WCuevocsNXU7fzNasO1JiMATzRKKZ5dII0KHfne93pYrC0fmkZH3qcL4DrNX3MkrgRqe9rKNJb+iBQw/EMP0J9k4Q7WGPBFPPao9DbHS24pPvc8hOaMz6Eo0YGq8PdVbfc0CcR8c9R3FA4tCrhl7VTVgid48DER72SDCpdLseV73JmlFcj17c6jXXOpI/bg8vFlFeDOFRkJkeQzivkGbte/70hzMC4Rb2SDC5MyaSP8ML3wR6EqKEeeZIyj0v3fYnkZgKZwRiKuxRp8cV5fYsbXcaCfCNyaI1R1JSQnYLOK8t64xA/IM96vQMlFy4ZVtG1AQJ11lePQIDGfyjNH3Lygh3yYYjS7RHfrLy41Uy8Ge1Vcyd8Lj/sJ7vlrXpEkT6Cl4LyBy1ScXujEDEgD3q9NSDOjSlCbt94xOfUGuOxM40VCA1gf9ujzqNcpYLqcpjo7KAwnSG0xz1cDfJe27TMYGbbY86DY6UDLfeSwYECMd27ssc6YlTbIe5yFkvxwSesj3qNJBKtn6S58qLd/HGX8e4h3TG2+qm3uyltJFX5btiVdajTzFsBXNeJ++yCeMge9vr2wvjyWW023rPL/b/vlv4sA+kSR2PbWakHy7hn6dX907p1Z9crLfBQ2i89u6e5U39ZRH57sXvhOxTqZ9kZTWeG2/ChwDrnUemKcHf4WtwVhvzNy2H7m1Cz9QUTHCMAV6taxMxjvGtTkwJx9K7QlpwVQL9SVPIn8RCd6SCafk3X5R7QlpZZ3tDhKIPWGK3vVQ5g+6hKeQe4uBM8Yk3pbCdAw2VJ2K1uWTFMaGjZvI6aulLF+7xAVhdXV6R8lum8c8P98/nK4uG1FVm1ybL5PR52ddouPUiK80iSPSYwjun5XtmJEV5cr3eRs/S7NWLMEK7qrRaJISnNSmtFWsbj2epCL0uk9fr6gC7FktArQTiWw9aLZLKS4FRcJEHjS6T1+jqgMtDsMR3iWtzunKwkw1KezoPb79L3v7u+PLaEW0nqEqhlNGwtaxa0O8yef0uTlbNr9wezd2Fty1Smal1aKdrG+bDljBCVlqu7XxtQfvJ5LWfumBN54QtiB2+OdqG2nrHBIpN0hb0h0xef6gLGPXP0oJIIp2Y7XbRfnLQbDQzQxfHFDwUuyhdkJsV5R0h4hisQ78cNymfTSOWcTfu8fManJusUHFeKOMJeUvT8AwH70wtputx3eJMaVnYOJxtyDOPQ4vGEjzWpotMPbYMP6QI53UpeT/U8So460P/xRI8H6YLXf0LAh6Asa4CpBzfV0dxowUC9Hgs8fsMq02b3Ac/gxyz7zLMS5ifmDr5O8VHyVkYqL/dkF0ZKJcSYrGqVXygaWPZuPGRPASi4QnroCmmaVb5gu6M5eDm65UfkVbdhtvsYtnec1Utj26gHLFb+6pAQVvFcnEDJW27gQYpfBsURyyotwKVDf0PS/C8T8MNQVLBWWQj547g9AcIdJbsZdAujRwpZ3z51/untuFtQw6dxYZ+iSV4XEcfm5i2qyxYka67nI4lQnF4m2xXbWi4WIIHcvTRLGll2EEsIi/76NaGrSIQbWiWWN2YJRLgwiFKs0BBuOLrWMwTHx/FRGkb+ik2r5/y1t+GERkekm9MBoukJW+i/d+eXj8mJTrGUJXgTT5gnD5Ag7cYpJ1GZx9bb9lO0bOaE1so4yYpok9K8Mx48kD2NcsvkxhyUpzt/uckq9OROrQLVHhqy3pv8YVL8LY5pZk5W5dncQqdFJvXSdENbZsBVze/CIdjKBMiQllwKIaeis17bkY31RfLe+/p2rvf/hCKLlL5oJa4SpXSFUCBdfcEF+qbzJ5QBQ+6MTbvERrd4LXe6vNFT2LDD6l64AoEYlYhhJaNzXugpgjf7Q9+tAqXxt8TJH/009S/lB3U4VdJiuXsD85BeOa0YPMBUCeAlMo6kNiDEQ/qrNywQTvH5j0lgwC2Dg8mNIQFyU7HhWomd2A4TzszmlZqUgoHukQ2r0uEgFU9BxsaYotvg+NCaRIXwWJ7HAfaSjavrYQA5l6vNkU6Atb1JVRfead8B1pAdtkC4ongOFVnilW2hSH0k+LmObM4IznOP6gvcYvnS4JxZTBuleLEgZaPPesDTJ3ktNaJF+IM1uNWGawdaNQ4o74hpzZXNRruEGamHrfKTO1Aw8Ux+wDfYR0Th8v9x3yZJZW500Vh65C8+HyZurjP75MlU5Ib/bRii8lCzl95LwydmFLrQzebk1ZRlA60TRzxoyga2OxRgK4gZfi20ROxmMwqYtAgccSPq2gZ/mRmbkbHGdJUz5N2ewtohDid3qKz8EqXjrKSiB+cv0MSaXl7ViGlj16GJy/JmwTL5PbH54fCoFWpJ3rud660M0fz8qjV0FlRSm9MK9ap8QeAPx2AOZpaDsMudFycTi+LkjACy9z4iF1+efhGjbdgtulSOHdIeq5MEZ8hhC7DdKEZpLSfKZorhq42dDU9c5lYb4MGmNPp7XISZjU9p1XaMI5PCpnquYJWfC4RiyxwoSWnFG9Fk8lAeF8IlzOEC0JOMQE7vURRys5E221y7RYrKLXGqZTz2p3sDITvknMprqNK2pXtDwbg+wa8rPW5OPMUN7PTq0MljPAqb06Sjjmqq5CmfXECirVWzzPFOlXJs6IxfEAaN9Jyxmt+qikerWDKD+2jdHJze2i8Dh4j+cEn11nUE2nwcJ0GwSXKY7hLYgCywHXyk5UfrxJ1MauZYrzA4/7Der5b0kMEWm1VK0XW3jdaRYS6FKtVJYSqFA85HHand/QbIr1qR7nuasfYMXRhXcFM//rHWC+6C4NF5G3VbPrKtlDx24xfGRdvb/viBIF60t6pxNJb61uIk0Y+35Ogz6q0J6maKIbONHQm1TMUV3+CTqoreJHHkzfhOiDfzupHV0E03wXbLAXIT4cUIMdek136fJ2urb3k/V97D6En0FmSPhDG22bHUzF1FHZNNO0Ts2Of8E5d4vp+uUt2OUlPuEhb8GJ1F5CfkM7B0x3Shz6xvsy7hGW/MKq/Pv4KdBbuA99suOg8sSsv7wZHTEAXgG6rK3THewddgBrD2Qr8NgSeDdXshDuHzV2LRAPyuMtqiY0d9D9d3kxVXWH38re/wcTd4XU0g+eywRM4068OvLya2ORBL9IVimtPlibBYpdlPC6vmb01KXp5VZz+6LgWK/3toWGvyQKFmmuKgdq4wdx+64eR4S3nYWJkzUMyvWff2v0MHxVqoCx2l18RAev2hMm6DU/hOu8dAdZLJJzAqePlxAkbbmg6ukKBuNLhbrZ8bYI4huUrE25ky1Z7yiYb7vGUk312kQq9R1coJFA61iJLUj6Ym1CFYinqmEyWKLfqKGfp3MITGn5uORTplR8lcJwOt5bOtUb+Kv3i1z55+V3kMYAybnbLVf73R7AOyxX4kUODpk9l5dPY/xt5K7nJ1dLSHF/rt4YJECu8ERMz0yJNXhGfTvpnsba4Zaf0kTlF5Vk4/U0BPzgR1yw6QdtS2xO2D7vlTvVTrzRBi8+dosGOtvRCwp/QqlAxiNNaEOG82SmIaUVdlCqqlsgJxZWboSGSeWxXC4Tn5lnF3M1quYPHnDrljuMY7gTaVuMRGtRuIoLH/sOalQCSi7iug2c4oC2XSxe75qgWXhMmu+fNAzmB5s3YVEDV+pePZgmpd97+wyK7vjN7RkgawJS7jDuTourRj+LE+3oMHpJPHTepNazkb52/KbY1nimQb15wFDu0H7vZONdwE2iVjC01wFnUMaz86hZ24k6jk4USOYGs8+LIWcVZk9JwvMxBn2Rsd8EctbegR+44yOEkTsgRbkQcu9l4gYP2yNiRCdwV+We0/3C/e/AM16RD5xmFT0lfqzXYHZxeLH9ljCstp36lJWQL88FHaUtmE5YqqR5BaGKMXTQIapHo6vnDMLTVi3SmApVOBD7ewQ/6DOMxGvL0qnT1CGKS6cx6nc5UINSJsCgs1EHHYizVsWgFZedCXT2HmpU6s16qMxVodSLo8Wh10KgYSzUqWtH2xotjz4j9Oa4FYOW1dNFn1WttloDWpoK+rJbe5bVUS+EUmhNjqeZEKwp/k191KZs/mbcFcVB7LIY2Xs16XpvfQTmFjsNEquPQip+r/d+Wu6UXG98Hjx7rBIymkSyYk1aeh5vAM/4ulREO7/p4fFdtuFj1uCjQ80SGN9iq9UMctC8mDW7iSN86Pr6Tv17sf37w15BYdhb6vETpBWJFTC+j4H0bOBvKe7Lv1airoRL7x/JqA1xAH6zSBL2JSYO7MxTSpM5m1YQLDtXPqlf92iAFrYdJg7syFCKlyIHQxBMCFa9yuAPAJCAhV2GCtsLEQQWTRHUNE1WYhDmrXpgTOT9UxQtaBhMXFV7tdDJMRGmW2Kx6ic0SUHerEEH1fzJGBZEc+QsTTEiUM/uMciag21ahgur9ZIIKqrZqFiactEtadr2kZTcPd5xCYX7S4NprhSC99JMrKzbJtihYw6d/LjLCuWrQhqaAPlZhbwbl+IlcOT65LpD8waoE2sUqjNO46ehY8aFR+ujZnSDXaX35cWin150XXxfbfs9WEJBWap/zLXiqBeaRvRkU86dyxfyW9Ike4Beiz689x//8cI7fN3br9PDkC0M5uZj2lLaKsDYxhDnO5c+g2D+VelahLcLNU+zoGTw1b0JtoVQnSqg7c855Bg2BqdTDCm2Jq93LIiYPy44VJovtmsByPAebROgjTKUeYWhLImMDzJXNQQ+D+re5UBBhJnao4gCdgKnUAwb5iYiLZXB6CrXktE1tcl7imGaiuULCl/JDerzP7W+T8nmlkifN7t09K5UaF0xOvfhGGem4WYO2wFTqSQL5rGk5XtCIGwT7V+eM2qbuaDxfi9MamvMowgyaEVNxM+JN/u9GGD2Q70l2jGsyZe82fnS//zAPfEA0W/b7lkxg0TrNzOcbd2FWWn/96MXlS7L+1bszLm5vjLj8Lfw4c1xLIVXa46qkEvrMMurqCK5d2xFcGEmZ5W0OuUZRaGFMxS0MmShevx9QRItifUCSy7rSkw9FaIJMVaRRkpnzQeVe9oVxG2zS1Kh57sJFQD76kHSGbeTniQvJnoIsLv2VR0UOdXaI4tkZ/Bki3BH0SaYqcipJzRDRCz6vkeaSKALag3wS7gh6KTMVqZhk5pPoBaBIM08U+cSffcIdQZ9khsonwWo0eye1EK4RkQBoK5jBpbvN7ggaJzNUxsmn6DbX4ovCbi5yrOKgmWzL2R1B22WGynbBaDnXcqjLcy6ipz9gp953dkfQ3pnJzR/1CfvOtfjpNp6LGOoPuuEzn90RdIBmcnNJfZLmcy2I+gzDIoIwAofLFXx9dTEaVQVt8sOnN/mPj2C89u7CyNvuf04uDCGvfOVlAbylWiM/MzaRv/GSfTmA4x2pRM8I75b7D/+yK3uAV96KlGP/f7PbDuf5sw9VvP+f60UwT8nLni7oCHJwkt+09A9G4btusu/6Ki2r8QT+5suvkosYA/Ksh4zuZOd3vDhlnkP7rLAoPNTW88zJPu6zg8I28Egmw9sGlVzCkVFXdDivLsCNGMljD69JnhCsvLT4TCbX/vvQ+O7dG+PJxgti78FfsTPAcyNpVpXDDElzhAbJb1ebXeT/SKoqWB8i/6UCWf2CzFM5Js/2knCIVQatsQhiwtP+w6O/pBotizxyohFQ1JJqwImSG5ufp6rOl/NkouHpu1W6gVQ3ruVf8NVxCEqL95WRvfFv8jc2TDIm1Q9hh0TsyS+NcnWoGM+KFaODOoE7ewB1Vfkup85CQ92r4C5KLlRSMH6dgIl94/Q9BJDkTuNNsp4yvPwXi6x5XpzYPIK09VZ3+7+tksLPScusklI8esnFSOkKtRFSlVLroApqJvxUVRW5nCobGVWEm2AVblWmxD5FEsTkAfsP94kqtzjgdniBF9nUmcyLO7I2itKB68RAC4RKRdSBUcN0JilGVUEsx8hBg9FlSCaZ/Ucy/stf7heeXVlYJScl/ThdM6UfInvq/CP3YZBAFAcLP7u9JBOHk2c3YqhaPh0AQeWCexyy6ABZTQBaLtIIywMBkvgRTtjFv55KQNgm81LyJck9NscyJAV8TvaH69gzvgkW0f6n3TLZ5fvZp8gv0wJ9ZZCJjbRlvsIK4+ISi5RzaZAKSubI/f8mZY+Chzy2UmhhValWkfReSrGjZM7hx66qqObYuViwu6G6UXKpu0lMpx5Cd6gbHcy1UMisqmyaMzfGwtzFW3Ubx4u3VYg4doWFUtI3hU3oyUupA56Ghy5SeKpBmjk8MzTwzMmK5kcj3i3v9x8TRZrCkgwnCHxPci119jO/wM/xXZ/HaYmON04nZMEx7jmt5hrRRa0GHbA1TK+TwGbThVO7iXCqBLbzeeUYqHHEnp9Qq0Zq/OoQqbFOD5QVZAY2d4px4w/jUAobJQUPbwB6yhtdWLWbCKtKeHvpvye7OmrKneagHR56nAmLujzHuow5cXqZUUraJysLYyhswlupInSA1vzQTQoaXUu1m2ipSkCTHL+z+ea4Sq/cmIoRrtpYILVYNUzjlDJFV1LtJkqqmq2k9RtX6sCVPLCqd8k0Emv3fofC6KAECqVCgw9d6rLRSF0XDJG0OSnpEwVHnmSUKYsFL/Kl+ybyH/f/Qd04lgMhGi3cdSqolNiverL+8f8DcZ+43A=="
VERSION_BASE = "Planos industriais + plano de cana incluído em 19/10/2026 · vínculos validados e pendentes"
REQUIRED = ["Grupo", "Material", "Equipamento", "Local", "Característica",
            "Tipo de Análise", "Dt. Resultado", "Vl. Resultado"]
KEY_FIELDS = REQUIRED[:6]
ITEM_COLS = ["ID", "ID plano", "Grupo", "Material", "Equipamento", "Local",
             "Característica", "Tipo de análise", "Previsto", "Realizado",
             "Faltantes", "Excedentes", "Atendidos", "Aderência (%)",
             "Condição", "Situação", "Observação", "Linhas do monitor",
             "Arquivo do plano", "Referência no plano"]
SUMMARY_COLS = ["Grupo", "Previsto", "Realizado", "Faltantes", "Excedentes",
                "Atendidos", "Aderência (%)", "Itens comparáveis",
                "Itens pendentes", "Registros no monitor", "Registros vinculados",
                "Registros pendentes", "Registros sem vínculo"]


def normalize(value) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = text.encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "", text)


def load_rules() -> list[dict]:
    return json.loads(zlib.decompress(base64.b64decode(BASE_REGRAS_COMPACTADA)))


def is_blank(value) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def parse_datetime(value, epoch):
    if is_blank(value):
        return None
    if isinstance(value, (datetime, date)):
        return pd.Timestamp(value).to_pydatetime().replace(tzinfo=None)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            converted = from_excel(value, epoch)
            return converted if isinstance(converted, datetime) else None
        except (ValueError, OverflowError):
            return None
    text = str(value).strip()
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y",
                "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def read_monitor(content: bytes) -> tuple[list[dict], list[dict], str]:
    """Lê uma única aba de monitor, detecta cabeçalho e preserva zeros/textos."""
    if len(content) > 30 * 1024 * 1024:
        raise ValueError("O arquivo excede 30 MB. Exporte um período menor.")
    try:
        wb = load_workbook(BytesIO(content), data_only=True, read_only=True)
    except Exception as exc:
        raise ValueError("Não foi possível abrir o Excel. Envie o monitor no formato .xlsx.") from exc
    try:
        matches = []
        required_names = {normalize(c) for c in REQUIRED}
        for sheet in wb:
            for rownum, row in enumerate(sheet.iter_rows(max_row=20, values_only=True), 1):
                names = [normalize(v) for v in row]
                if required_names.issubset(names):
                    if any(names.count(name) > 1 for name in required_names):
                        raise ValueError("Há colunas obrigatórias duplicadas no cabeçalho.")
                    matches.append((sheet, rownum, names))
                    break
        if not matches:
            raise ValueError("Cabeçalho não encontrado. Colunas necessárias: " + ", ".join(REQUIRED))
        if len(matches) != 1:
            raise ValueError("Há mais de uma aba de monitor. Envie uma exportação com apenas uma aba de resultados.")
        sheet, header_row, names = matches[0]
        positions = {c: names.index(normalize(c)) for c in REQUIRED}
        extra = {c: names.index(normalize(c)) for c in ["Dt. Registro", "Usuário"] if normalize(c) in names}
        records, invalid = [], []
        for rownum, row in enumerate(sheet.iter_rows(min_row=header_row + 1, values_only=True), header_row + 1):
            if all(is_blank(v) for v in row):
                continue
            rec = {c: row[i] if i < len(row) else None for c, i in {**positions, **extra}.items()}
            rec["Linha na origem"] = rownum
            dt = parse_datetime(rec["Dt. Resultado"], wb.epoch)
            missing = [c for c in KEY_FIELDS if is_blank(rec[c])]
            reasons = []
            if missing:
                reasons.append("Identificação incompleta: " + ", ".join(missing))
            if dt is None:
                reasons.append("Data do resultado inválida ou vazia")
            if is_blank(rec["Vl. Resultado"]):
                reasons.append("Resultado vazio (não conta como realizado)")
            if reasons:
                invalid.append({**rec, "Motivo": "; ".join(reasons)})
                continue
            for field in KEY_FIELDS:
                rec[field] = str(rec[field]).strip()
            rec["Dt. Resultado"] = dt
            rec["Data"] = dt.date()
            rec["_key"] = tuple(normalize(rec[f]) for f in KEY_FIELDS)
            records.append(rec)
        return records, invalid, sheet.title
    finally:
        wb.close()


def rule_keys(rule: dict) -> set[tuple]:
    base = tuple(normalize(rule[k]) for k in
                 ["group", "material", "equipment", "local", "test", "sample_type"])
    keys = {base}
    keys.update(tuple(normalize(v) for v in key) for key in rule.get("observed_keys", []))
    return keys


def sugar_product(value):
    """Identifica a classe sem confundir VVHP com VHP."""
    value = normalize(value)
    if "vvhp" in value:
        return "VVHP"
    return "VHP" if "vhp" in value else None


def select_sugar_rules(rules, day):
    products = {sugar_product(r["Material"]) for r in day
                if normalize(r["Grupo"]) == normalize("Produção de Açúcar")}
    products.discard(None)
    active = next(iter(products)) if len(products) == 1 else None
    selected, excluded = [], []
    for original in rules:
        rule = dict(original)
        product = sugar_product(rule.get("plan_material", ""))
        if product:
            rule["group"] = "Produção de Açúcar"
            if active and product != active:
                excluded.append({"ID": rule["id"], "Produto": product,
                    "Ensaio": rule["plan_test"], "Ponto": rule["local"],
                    "Previsto original (24h)": rule["expected"],
                    "Situação": "Não aplicável", "Motivo": f"Produção exclusiva de {active} no dia"})
                continue
            if not active:
                rule["status"] = "Pendente"
                rule["notes"] = "Conferir produto do dia: ambos ou nenhum identificado em Produção de Açúcar. " + rule["notes"]
        selected.append(rule)
    return selected, excluded, active, sorted(products)


def evaluate(records: list[dict], selected_date: date, rules=None) -> dict:
    """Apura cada obrigação antes de somar. Nunca compensa faltas com excedentes."""
    rules = load_rules() if rules is None else rules
    day = [r for r in records if r["Data"] == selected_date]
    rules, not_applicable, active_sugar, detected_sugars = select_sugar_rules(rules, day)
    by_key = defaultdict(list)
    for rec in day:
        by_key[rec["_key"]].append(rec)
    prepared = []
    owners = defaultdict(list)
    for rule in rules:
        matched = {r["Linha na origem"]: r for key in rule_keys(rule) for r in by_key.get(key, [])}
        match = list(matched.values())
        counts = Counter((r["_key"], r["Dt. Resultado"]) for r in match)
        status = rule["status"]
        note = rule["notes"]
        if status == "Comparável" and any(n > 1 for n in counts.values()):
            status = "Pendente"
            note = "Chave/horário repetido no monitor. Confirmar correção, reanálise ou amostra distinta."
        item = {"rule": rule, "match": match, "count": len(counts), "status": status, "note": note}
        prepared.append(item)
        if status == "Comparável":
            for rec in match:
                owners[rec["Linha na origem"]].append(item)
    for items in owners.values():
        if len(items) > 1:
            for item in items:
                item["status"] = "Pendente"
                item["note"] = "Registro associado a mais de uma obrigação. Rever vínculo antes do cálculo."

    rows, associations = [], defaultdict(list)
    for item in prepared:
        rule, actual, status = item["rule"], item["count"], item["status"]
        planned = int(rule["expected"])
        comparable = status == "Comparável"
        fulfilled = min(planned, actual) if comparable else None
        rows.append({
            "ID": rule["id"], "ID plano": rule["plan_id"], "Grupo": rule["group"],
            "Material": rule["material"], "Equipamento": rule["equipment"], "Local": rule["local"],
            "Característica": rule["test"], "Tipo de análise": rule["sample_type"],
            "Previsto": planned, "Realizado": actual,
            "Faltantes": max(planned-actual, 0) if comparable else None,
            "Excedentes": max(actual-planned, 0) if comparable else None,
            "Atendidos": fulfilled,
            "Aderência (%)": fulfilled/planned*100 if comparable and planned else None,
            "Condição": status,
            "Situação": ("A menos" if actual < planned else "A mais" if actual > planned
                         else "Na quantidade prevista") if comparable else "Pendente de validação",
            "Observação": item["note"],
            "Linhas do monitor": ", ".join(str(r["Linha na origem"]) for r in item["match"]),
            "Arquivo do plano": rule["source_file"], "Referência no plano": rule["source_rows"],
        })
        for rec in item["match"]:
            associations[rec["Linha na origem"]].append((rule["id"], status))

    linked_rows = []
    for rec in day:
        links = associations[rec["Linha na origem"]]
        state = "Comparável" if any(s == "Comparável" for _, s in links) else "Pendente" if links else "Sem vínculo no recorte"
        linked_rows.append({**{k: v for k, v in rec.items() if not k.startswith("_")},
                            "IDs das combinações": "; ".join(i for i, _ in links), "Vínculo": state})
    detail = pd.DataFrame(rows, columns=ITEM_COLS)
    linked = pd.DataFrame(linked_rows)
    groups = sorted(({r["Grupo"] for r in day} | {r["group"] for r in rules}) - {"Sem grupo confirmado"})
    summaries = []
    for group in groups:
        items = detail[detail["Grupo"] == group]
        ok = items[items["Condição"] == "Comparável"]
        recs = [r for r in linked_rows if r["Grupo"] == group]
        n = int(ok["Previsto"].sum())
        summaries.append({"Grupo": group,
            **{col: int(ok[col].sum()) if n else None for col in
               ["Previsto", "Realizado", "Faltantes", "Excedentes", "Atendidos"]},
            "Aderência (%)": float(ok["Atendidos"].sum())/n*100 if n else None,
            "Itens comparáveis": len(ok), "Itens pendentes": int((items["Condição"] == "Pendente").sum()),
            "Registros no monitor": len(recs),
            "Registros vinculados": sum(r["Vínculo"] == "Comparável" for r in recs),
            "Registros pendentes": sum(r["Vínculo"] == "Pendente" for r in recs),
            "Registros sem vínculo": sum(r["Vínculo"] == "Sem vínculo no recorte" for r in recs)})
    summary = pd.DataFrame(summaries, columns=SUMMARY_COLS)
    return {"summary": summary, "detail": detail, "linked": linked,
            "no_group": detail[detail["Grupo"] == "Sem grupo confirmado"], "date": selected_date,
            "not_applicable": pd.DataFrame(not_applicable), "active_sugar": active_sugar,
            "detected_sugars": detected_sugars}


def total_row(summary):
    result = {"Grupo": "TOTAL PARCIAL"}
    for col in SUMMARY_COLS[1:]:
        if col != "Aderência (%)":
            result[col] = int(summary[col].sum())
    result["Aderência (%)"] = result["Atendidos"]/result["Previsto"]*100 if result["Previsto"] else None
    return result


def csv_bytes(frame):
    # Evita que textos de um arquivo enviado sejam interpretados como fórmulas.
    safe = frame.copy()
    for col in safe.columns:
        safe[col] = safe[col].map(lambda v: "'" + v if isinstance(v, str) and v.startswith(("=", "+", "-", "@")) else v)
    return safe.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")


def percent(value):
    return "—" if value is None or pd.isna(value) else f"{value:.1f}%".replace(".", ",")


CSS = """
<style>
/* Paleta do aplicativo de referência Carta controle NIR DS-3. */
[data-testid="stAppViewContainer"] {background:#e8edf3; color:#083050; color-scheme:light;}
[data-testid="stHeader"] {background:#e8edf3;}
[data-testid="stSidebar"] {background:#f5f7fa; border-right:1px solid #dce3eb; color:#083050;}
.block-container {padding-top:4.5rem; padding-bottom:2.2rem; max-width:1600px;}
.qmp-header {background:linear-gradient(118deg,#00983a 0%,#004881 64%,#083050 100%);
 border-radius:12px; padding:24px 34px 20px; margin-bottom:6px; box-shadow:0 3px 10px #08305012;}
.qmp-header h1 {color:#fff !important; font-size:31px; font-weight:750; line-height:1.25; padding:0; margin:0;}
.qmp-header p {color:#eaf5fa; font-size:14px; margin:8px 0 0;}
h2,h3 {color:#083050 !important;}
[data-testid="stAlert"] {border-radius:9px;}
[data-testid="stExpander"] {background:#f5f7fa; border-radius:12px;}
[data-testid="stFileUploaderDropzone"] {background:#fff; border:1px dashed #bccbd8; border-radius:10px;}
[data-testid="stDownloadButton"] button {background:#004881; color:#fff; border:1px solid #004881; border-radius:8px;}
[data-testid="stDownloadButton"] button:hover {background:#083050; color:#fff; border-color:#083050;}
[data-testid="stDownloadButton"] button:focus-visible {outline:3px solid #00983a; outline-offset:3px;}
[data-baseweb="tab"] {color:#083050;}
[data-baseweb="tab"][aria-selected="true"] {color:#007c30; font-weight:700;}
[data-baseweb="tab-highlight"] {background:#00983a;}
.group-grid {display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:16px; margin:20px 0 28px;}
.group-card {background:#fff; color:#083050; border:1px solid #d7e0e9;
 border-top:4px solid #004881; border-radius:12px; padding:20px 16px; text-align:center;
 min-height:185px; box-shadow:0 3px 12px #0830500b;}
.group-card.complete {border-top-color:#00983a;}
.group-card.complete .group-value {color:#007c30;}
.group-card.empty {border-top-color:#9aaaba; background:#f5f7fa;}
.group-card.empty .group-value {color:#64778a;}
.group-title {font-weight:650; min-height:46px; font-size:15px;}
.group-value {color:#004881; font-size:36px; font-weight:750; line-height:1.3; margin:7px 0;}
.group-note {color:#506579; font-size:12px; line-height:1.7;}
@media(max-width:1050px){.group-grid{grid-template-columns:repeat(3,minmax(0,1fr));}}
@media(max-width:760px){.group-grid{grid-template-columns:repeat(2,minmax(0,1fr));}
 .qmp-header{padding:22px;} .qmp-header h1{font-size:26px;}}
@media(max-width:440px){.group-grid{grid-template-columns:1fr;}}
</style>
"""


def style_summary(table):
    """Apresentação da tabela; não altera os valores calculados."""
    def row_colors(row):
        if row["Grupo"] == "TOTAL PARCIAL":
            return ["background-color:#083050;color:#ffffff;font-weight:bold"] * len(row)
        base = "background-color:#ffffff;color:#083050"
        styles = [base] * len(row)
        for column in ("Faltantes", "Excedentes", "Aderência (%)"):
            value = row[column]
            if pd.isna(value):
                color = "background-color:#edf1f5;color:#64778a"
            elif column == "Faltantes" and value > 0:
                color = "background-color:#fff3de;color:#805200"
            elif column == "Excedentes" and value > 0:
                color = "background-color:#e3eff9;color:#004881"
            elif column == "Aderência (%)" and value >= 100:
                color = "background-color:#e2f3e8;color:#007c30;font-weight:bold"
            else:
                color = base
            styles[row.index.get_loc(column)] = color
        return styles
    return table.style.apply(row_colors, axis=1)


def cards_html(summary):
    cards = []
    for row in summary.to_dict("records"):
        value = row["Aderência (%)"]
        empty = pd.isna(value)
        cls = "empty" if empty else "complete" if value >= 100 else ""
        note = ("Vínculos aguardando validação" if row["Itens pendentes"] else "Sem base diária vinculada") if empty else (
            f"{int(row['Realizado'])} realizados / {int(row['Previsto'])} previstos<br>"
            f"{int(row['Faltantes'])} faltantes · {int(row['Excedentes'])} excedentes")
        cards.append(f'<div class="group-card {cls}"><div class="group-title">{escape(row["Grupo"])}</div>'
                     f'<div class="group-value">{percent(value)}</div><div class="group-note">{note}<br>'
                     f'{int(row["Itens comparáveis"])} itens comparáveis · {int(row["Itens pendentes"])} pendentes'
                     '</div></div>')
    return '<div class="group-grid">' + ''.join(cards) + '</div>'


def main(monitor_bytes=None):
    import streamlit as st

    st.set_page_config(page_title="Aderência Analítica", page_icon="📋", layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown('<header class="qmp-header"><h1>Aderência Analítica do Laboratório</h1>'
                '<p>Previsto × realizado · Rotina diária · 24 horas</p></header>', unsafe_allow_html=True)
    with st.sidebar:
        st.header("Monitor de digitação")
        uploaded = st.file_uploader("Importar monitor (.xlsx)", type=["xlsx"])
        st.caption("A base dos planos já está incorporada ao aplicativo.")
    with st.expander("Plano diário — Laboratório de Cana"):
        cana = [r for r in load_rules() if r["group"] == "Laboratório de Cana"]
        st.caption("15 itens · 61 resultados previstos em 24 horas. Vínculos com o monitor aguardam validação; não entram no percentual até serem confirmados.")
        st.dataframe(pd.DataFrame([{"Material do plano": r["plan_material"],
            "Ensaio": r["plan_test"], "Frequência": r["frequency"],
            "Previsto em 24h": r["expected"], "Coleta": r["sample_type"],
            "Validação necessária": r["notes"]} for r in cana]), hide_index=True, width="stretch")
        st.caption("Fora do total diário fixo: análises por carga amostrada, por frente de colheita e uma vez na safra. Comparações NIR não são automaticamente resultados da rotina.")
    content = uploaded.getvalue() if uploaded is not None else monitor_bytes
    if content is None:
        st.info("Importe o monitor de digitação para visualizar os cartões e a aderência por grupo.")
        with st.expander("Critérios da avaliação"):
            st.write("Frequências diárias: 2h = 12; 4h = 6; 8h = 3; diária = 1, além de 2 ou 3 vezes ao dia.")
            st.write("A base contém 300 combinações comparáveis e 221 pendentes. O indicador é parcial.")
        return
    try:
        cached_read = st.cache_data(show_spinner=False)(read_monitor)
        records, invalid, sheet_name = cached_read(content)
    except ValueError as exc:
        st.error(str(exc))
        return
    if not records:
        st.error("Nenhum resultado com identificação, valor e data válidos foi encontrado.")
        if invalid:
            st.dataframe(pd.DataFrame(invalid), hide_index=True, width="stretch")
        return
    dates = sorted({r["Data"] for r in records}, reverse=True)
    with st.sidebar:
        selected_date = st.selectbox("Data do resultado", dates, format_func=lambda d:d.strftime("%d/%m/%Y"))
    result = evaluate(records, selected_date)
    summary, detail, linked = result["summary"], result["detail"], result["linked"]
    if result["active_sugar"]:
        st.success(f"Produção de açúcar em {selected_date:%d/%m/%Y}: {result['active_sugar']}. "
                   "O outro produto está como Não aplicável, fora do previsto e dos faltantes.")
    else:
        st.warning("Conferir produção de açúcar: " +
                   ("VHP e VVHP encontrados no mesmo dia." if result["detected_sugars"] else
                    "nenhum VHP/VVHP identificado no grupo Produção de Açúcar.") +
                   " As obrigações específicas dos produtos ficam pendentes, fora do percentual, até a conferência.")
    if not result["not_applicable"].empty:
        with st.expander("Análises não aplicáveis — produto não produzido"):
            st.dataframe(result["not_applicable"], hide_index=True, width="stretch")
    with st.sidebar:
        selected_group = st.selectbox("Grupo", ["Todos"] + summary["Grupo"].tolist())
        st.caption(f"Data analisada: {selected_date:%d/%m/%Y}\n\nOperação considerada: 00h às 24h")
    filtered = summary if selected_group == "Todos" else summary[summary["Grupo"] == selected_group]
    filtered_detail = detail if selected_group == "Todos" else detail[detail["Grupo"] == selected_group]
    filtered_linked = linked if selected_group == "Todos" else linked[linked["Grupo"] == selected_group]
    total = total_row(filtered)
    st.info(f"Aderência parcial: {percent(total['Aderência (%)'])}. "
            "Calculada somente nos itens comparáveis; excedentes não compensam faltantes.")
    st.markdown(cards_html(filtered), unsafe_allow_html=True)
    st.subheader("Previsto e realizado por grupo")
    display_cols = ["Grupo", "Previsto", "Realizado", "Faltantes", "Excedentes", "Aderência (%)"]
    table = pd.concat([filtered, pd.DataFrame([total])], ignore_index=True)
    configurations = {col: st.column_config.NumberColumn(col, format="%d") for col in display_cols[1:-1]}
    configurations["Aderência (%)"] = st.column_config.NumberColumn("Aderência parcial (%)", format="%.1f%%")
    st.dataframe(style_summary(table[display_cols]), hide_index=True, width="stretch", height=min(620, 38*(len(table)+1)),
                 column_config=configurations)
    st.caption("Células vazias indicam grupo sem base diária vinculada. Mesmo 100% se refere apenas aos itens comparáveis.")
    st.download_button("Baixar tabela por grupo (CSV)", csv_bytes(table),
                       file_name=f"aderencia_grupos_{selected_date:%Y-%m-%d}.csv", mime="text/csv")
    tabs = st.tabs(["Análises por ponto", "Vínculos do monitor", "Pendências e critérios"])
    with tabs[0]:
        situation = st.selectbox("Situação", ["Todas", "A menos", "A mais", "Na quantidade prevista", "Pendente de validação"])
        shown = filtered_detail if situation == "Todas" else filtered_detail[filtered_detail["Situação"] == situation]
        st.dataframe(shown, hide_index=True, width="stretch")
        st.download_button("Baixar análises por ponto (CSV)", csv_bytes(shown),
                           file_name=f"aderencia_itens_{selected_date:%Y-%m-%d}.csv", mime="text/csv")
    with tabs[1]:
        st.caption("Cada registro conserva a linha do Excel original e os IDs dos itens relacionados.")
        state = st.selectbox("Vínculo", ["Todos", "Comparável", "Pendente", "Sem vínculo no recorte"])
        shown = filtered_linked if state == "Todos" else filtered_linked[filtered_linked["Vínculo"] == state]
        st.dataframe(shown, hide_index=True, width="stretch")
        st.download_button("Baixar vínculos do monitor (CSV)", csv_bytes(shown),
                           file_name=f"vinculos_monitor_{selected_date:%Y-%m-%d}.csv", mime="text/csv")
    with tabs[2]:
        coverage = filtered[["Grupo", "Itens comparáveis", "Itens pendentes", "Registros no monitor",
                            "Registros vinculados", "Registros pendentes", "Registros sem vínculo"]]
        st.dataframe(coverage, hide_index=True, width="stretch")
        st.caption(f"Existem ainda {len(result['no_group'])} combinações do plano sem grupo confirmado, fora dos percentuais por grupo.")
        if invalid:
            st.warning(f"{len(invalid)} linhas do arquivo foram excluídas por dados incompletos ou inválidos. Confira antes de concluir a avaliação.")
            st.dataframe(pd.DataFrame(invalid), hide_index=True, width="stretch")
            st.download_button("Baixar linhas inválidas (CSV)", csv_bytes(pd.DataFrame(invalid)),
                               file_name="linhas_invalidas.csv", mime="text/csv")
        st.markdown("""
- O previsto vem da frequência diária do plano, por ensaio, equipamento, local e tipo de coleta.
- A data de referência é **Dt. Resultado**. O fechamento é de 00h às 24h.
- Aderência = soma de **mínimo(previsto, realizado)** por item ÷ soma dos previstos comparáveis.
- VHP e VVHP são exclusivos por dia: a identificação usa Material no grupo Produção de Açúcar; expedição não define a produção. O produto inativo fica Não aplicável. Ambos ou nenhum exigem conferência.
- As regras e os pontos conhecidos permanecem previstos mesmo sem resultado no dia, exceto o produto de açúcar identificado como não produzido. Pressupõe-se 24 horas de operação.
- Faltantes significam registros não encontrados na exportação, não prova de análise não executada.
- Tipos de coleta diferentes, equivalências técnicas e versões divergentes ficam pendentes.
- Mesma identificação e horário repetidos exigem validação; o item sai do subtotal comparável.
- Resultado zero conta. Resultado vazio não conta. Não há avaliação dos valores contra limites de especificação.
- Pontos novos e registros sem vínculo não são automaticamente excedentes nem entram no previsto.
- Paradas, mudanças de versão, novos equipamentos e equivalências precisam de atualização da base antes de uma avaliação global.
- Os registros do monitor podem incluir cálculos e análises fora da rotina diária. O total do arquivo não é o denominador da aderência.
""")
        st.caption(f"{VERSION_BASE} · Aba importada: {sheet_name}")
    if invalid:
        st.warning("Há linhas inválidas no arquivo. Consulte a aba Pendências e critérios.")


if __name__ == "__main__":
    main()
