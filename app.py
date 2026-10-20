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

BASE_REGRAS_COMPACTADA = "eJztnd2O3DiW519FW8AALkzZDn3Fxxh9kU7bVYW1u3PsrEKjuwcNZYQySz0RoRwpIsdd2wtsP8A+wN5tzV4MegADAxT2xpcVb7JPstRHREg6pERKpHhkC2h0VWVGKkTyd0ie8z88/P1/+yJYffEPxheXk8nE/OIr44v7tbf9Y/azq/xHd1G4v09+8PL9LvIO/374P6GxCo1Lb70Kk99vvJ0fBd46+chz7458IjSmv3w0rv1om37A/5d9cL/xt7vkE29Cf7vyjIvkF+twmf3ZO+/wgfxwVfnDnR+nf3MVrpP/jL3N/dr/4+7P937y02+38c7b7g7/tvW99Gve3/vLnZ+8+zT59M7b7eO0ceHm3osOPz346WO2IXlu2h7yXg/e+vAf22VAvtw3luHa33nkH9vbINp45I3u/XVo7OP94acoCA1/Y5izp+bkqTWxps+M24i07PjXoZF0XWhsyCsFK+/JqTNZ3bM7tjL9VLWp4T5a+n+8DdZpUy9+/eTdm2+fvP72108mFvmf8di4Ii36t42/i8I4+fKrKFz6cRwmzaCM0x9fh38O10/er+P3hcdH4b/GlMeb/2UdbH/wSFuTz6at9LfLPyefdN47PyQ/DG9iP3rwV3/8Z//PySN+/3sWHVQiihAwxj7viGTsQjKUX/zTP/33r4wirlafuD4fcR1xfd4FVxviauGYXb/bEABW/ueCbLG5KLCda5xlC53BRtfpE12hmXZE99NGt27G5ULXhejaOGbdi7efC7V5S1EAu9A412b9wGZ12ierQtPsyOqnyGrd5NrE6gyy6oiymv7QMNvNqn72TKP6hONAPo+C92LQmlzQUrk6t4RO1ellMHBlTQBXL4LEwDwxtCjDR5sEWQN17JPSiADO5v1xRp8RR87wcvZcHmcLyJmLaj4TjikpxAxT0McyUcxmeZfUQmZO+oOs3WQ2QqYLMq6pjAsyioozxTWT7SP/Rw8PZ+fXQYGahWM+O/VKPW0UEUYVbS2ntJE2rbTxTWyctFE0lFk72lrG8goNoAZKNHgD7FgGLm/AlklaY+iNNVBc3oBJETxUcdY4q42cIeOMNaO14IyiTsxRzWf9OwTDSDWwpihmMy6HgCIrqIKs3WQ2QqYLMq6pjAsyih6wwDWT6fADajhD5gfA1Ckt8xmnH0BRBVTR1nJKG2nTShvfxMZJG9QGskiuCG2Hv97t01d57T14d/4m7c545weRxyAvLoH3cku+gjQG/nlc5O7+GzHmbC7m+hfq8+7yDD9vN/nS9bnpfqHpJdDz9qOAHOZGzd/PBeX7BmwKpDQAknUMU8+3oDRhCidT98T4pRfdhMEq8nZhrIL24vONXxkXb68NQlyTCZDPIrEB0D8orAEms2i0BjKmZXP4yujtS2vXGguqN2Y51fYdefBdtL8PC+gxDNEz4tMKSAHHePTmD19c/OELQjf5l+d/+OJLYKUXRrAi/xrcBksvYmyEDj/l/UP7CtIzxibrIeOCfNFzFYZ8RZ5P3tIvW3HeO5ERJ+t/waSXYRT58X1I/sjwUvMmlhgH5JU33jPjKtzuQmObUHBu+yp8duruyCD/7cXEWEh7wns/Z6behFsMBH6btmGCB9Omq5xD3ci0VXJ+qQjvnG7jcgS7oecHwDNU+bl5hsqUWc5TuyZt9RLwmjzSN2TAQuPV4aebiAwQYPba25LXM672600+OABa4v6EmxuvSGSNXkAwJDTK90Y3oB2iooFNhqZ+5Cmdmoy8xzXy6eMj/yGIiRnkCMA9e71jyhhV1kjCwaOMGZd0YEGJynQRAceO6WrirTmwqwE3uCnuHzee8K4FhSpziom2unCbLuB4Ym4amDNF03KVQMcXerOgdmXOEHHHDLcxmOMLQLRnrjH8pYM3gV2VGtiyXqkHDepX5rwLaC/IiARrL7eqIe/gyi0ZwB7OFM1p42Cu3AkSt3FQyTIXqLjTtZGrwQ7lVs4UzTtSQx3Xbg4KWtYEF3T69nN13GHd0Tk40OPb09lQarJMVPRp2tXVkIdyX+dK3te1QI5jZ2dDRcUSV1RehyuYIcIOH196wXsvHdZ9RJ4WZQHS/BnHQX13+HkdrJgRYi7YqBFiqZHeY8tLPJZeHReVAqp7lRSoSVjimoQKUsT3/igoQeoeCCjRVUJglN9yUBDSbcnSxAfGVY1yaJ2bDhiSt1zVdLzwl2TpS7gw3r39djGuL6r5aK8R2jCIbk018vHC38Z11dMwA1J+d1yEwPATNyEw3G0Jh7uvw2iXrhGvgjVpMGAl+3FsPPqadPv6S4ZTlCcuGemHSq45NR50yqSS5hmdmnGbvq/x6O70uviDQZRDaPPfivpGlIGkjF3NkDWVJLVh0NsSDnorxo1aeUwfbY3lwXTABsM/4o54Z9gaCobZMNBtCQe6FbPGLiaqDzi+ip86qKOEf/qf4nhqgNow3G2Xw908+y4wYMF2FTwEq33+kpz7sW9Pf5UkiCe7oPyBzWssd+xb/Z6sqTMGsEKLna2sIOXAGLbdMoad8u5R9vR5APR0eqn4wRNO1/ub/fqY+0j7pKwQUouo9m3hNQYQFxI7oVAbzy4OQN0wNo4el4bswAi3baGGUXh+U8ciyslJJLG+LxR5hGUHRtBtGzeJLYRmhTAiFZhFEuN745FPbXZgyN5umZjfE5K9q89MHDHG521KdE3Y7ZCMIocK7UBtwG6Zrt8ThjLUA5kw4pUEbBnBPslIHnvLaCYTqhK2uCqRvQFhzPdWNVnWTSonaQt56eQYmbwZUb03zG48/umUEjtksVsFB4oV9qwjODWpXJ8DOgPLALMFQoBVeKD0YLfMt79MOiwbA/bEc2IE/E3d6hifAo/L498vswEr/31DkEV+9Piy+hZDO2dpw/Bb2zWTNf71Q95ypI+9yQw2O1DnsFsm9CMCuy/RV4BrnIEbWNhqEFg3KcUOlFCclicGMFHNCv7oBRtpEMihBCUHwfapP5l4u1DOcVrKOYjwvliS/emP/HjzHaBh14LLHkr2vEHB31RYI07ApgpdgcumRO8VwWNW5y6tDS+4UJxyWopTiEwrUfGXvqHGwoy/GL/2tz/sN4SWuyD1I1dJRR3D38YeMRxiF/epD1p0PzubCGwSLksRO9RbpRAKU05LYQoRhdTgRqe9yyAnd4yhEUdeTLjXWZ1DynChoua0VNQQ2dL1ProZt0u0zsBlV1KPOvdqWsVOrTcwqBU64ueI6oe7++kRcYPRFFsfOPACp6yrHEFlz2lZtAvRRE1MJ1wfPubz7Oe2+WHmiVe6BRfDAw32l/u0fs6GYqjTslAZIlsTOXwzzO0Qq7PRH/pxRK/QwGNWWXfWmxOUh53hy8N5AfhP2Z74orF5P+AyKKmlf3s2qOaLBVyoSzsSdOmahJ1PPuVicIURXUoat4R9WLkf9KRduFCgdiUI1Nrx1pd4MbT6i668xPCe4W5KvphCddqVoE7rZ1tn+sUAyzy6ajSFPghvTsGYQp3YlaATa4dcvkanVB4eWAay2z59fQolYVeCJKwdOFVCli7s8IbmXRiaF0hJmEIZ1ZUgo2rnT118vicCBxlidwVC7FUQodzotjyaiArETz94zTmDYgxfu8rC1z1sZzkC2FOovboStFcERvVZhbBrrQphENtVFsTuxayaw9hTqLO6wjprdnllVmjMeBsSQoMHeJz+ZX6bL+VzJ2u62BBC06rMxiXZehw+bPclc2m4CFRiugLl0mrlt7OTT+3Xu+BP5DvPd1STr8+KgnW6mFqH7cio/8REq4YmFkTgkuqqLUCR1BUWSflsIa/AN5oCjykc57vBWsJURvkpJlhsllobAtQ23bK2KXCZdOEiY8/Y+et0cPMyh/dR4moC62DnT5ZWwtPjTlwUHijTTNBfG83Xx0OyGEponjNLcwqVy6l4dVIc8NbUDEKSE9weQLz1h6YC9Z8r8M2gtDgtS4unxFPy5hekXz/mlEC/6So6/O3xywfvPqRWpyr/uoY08j3rIPmr05Xucfps4/zHflykrrcUkLQJfrGFokkfDhm+ChxestYt0+0zvMKeDwrnfIW9wO1Hv2eOLWs8KUPYbuQa0zlmUA2cWkPjsq/cjQYsm7M1lFNJ8azYwrU2KpvyMGZQMZzag4Oyv6SLJi550iyUo0m5DR0lm80ZFDMoKE6doeGJ/kR9A9Rcx+fVQy1WVqIKEhQEp+7QQMrLGawGUKShASlqS3TC1akSwwwKY9Pp0OASSerChlNjGpd6gMRuSq8CBBWg6awXgIztLx8nZh1GWeSXwzctaVyKNlzG+U4WlK4q5dZydb7qeeyYQ8Z118YMai590meN9A2WPksCfVDo6JO+6UjfYOmbSqCPolT0SN98pG+w9M270zenSBVzYfp+60XhfVopHmzu6dl02R+U0DtfR/g83NzwQMdwDabtgatpBz7cYJ68894Rxa2mxZTBooxRo+gwp4gOC2yE9SUq8AKGQF0wKbnv4tNZZ76a5IM5lA9mE3R49ScPcBOGQyegVHbXAVmzDjCHOsBMXNNXzFmfpUtVpf7xAqw/4mYJJfQpg5ejuOgcag8zceFfMbzfB/EyjAXvCu96GEI3xdVG68QZylvme2PFPOSgDOhKn9STDYWPmXj2gGKy+74yQDfUOBQ1CypqVv/+EWdB/znUX2biaQaKOR6Q+suLKjYZ2IIqHgvaKkFQQ5mJ5xeongl7PD2sfRZsOiCsHibRst7qZsHms71zqMLMxFMYlPPb40Fd/QA3ncVVT7BYMRClBDcfo51DKWcmLuXwIvzid4/nk8mkhK4C/UZ+ON14dFZyYD0QDLF1C8Y+5cbWC2PXTb9ZQP1GIXK/e+yMxA2TuPPQdQQOyjmzsmAocLAuP9/7uHLW3ThpvpHx7qVhmvmbc56wuwriZbD1zmeHizg2nATlFrLRnwYV7Nu2x0LVW4aARFBllaINtT7BXOrP02H0EVXZqDZ2LV5SBY4JVkmF6tJcXMV848UxGTHo7IQ/+tVU7fxnBnlXIylC8+BHhheRJ225KoEn8BHyCIAAPqv98p014A9fXPzhCzDSHKu0Szkl7BmvPNBx/MNLHml9eTwF+syofhvzfKjVJnJeGL/CkDWNFN+yDRWgubh8qQIweupF/3w1p1woxmvBxAse9OwRr7xf6umCKsxcXF9UQhcr80IDYDwZF4oZMydMyChHNvuk7NQ59aBBmWQuLvdlr/h8+Ovkc5TrpGmxIaNocu0he65kpYQ6ylxciVOBmI6VkkaY9pXStNmAUfSz/gDjWiuhzjEX1+mU8KVnraQihmGtdNiUwWhdn5RxrpVQjZiLC2pvvLsNFbSjapJ/wCghJ57Mr4qx5OXQLpTsLb8pbc9/Gj7qiNVk9NfBZU6g7jAX1x2kw9XvEslmS/sSabE3+pS01b7Q4lgfzQkUGObiJ5Lkk9X34lgDF4LF0TLZfME80t744loZzQnUBebiR5Le+GvjwngbLEPj1e8eUwP/2XvTP4hvrSSvmUYujKjyqsiWTYvtX9JyPluxRx2yuuFsvZDCyP+iReRfPYu9Lq18KOpfZdmOKCVzSR+JXOsuVAgWLRSCHkDseSXmZBHDosz2WCn5JRpx5FymoaqwaKEqpG91Fd5EfgOSyf6C/mE1a3WHU55nKO+r74ptsXbZTFL0cvEikYwhaxrS1us1VCAWLRSInqAUXrR7YFL/qs0W8i0pNXXlIcm1cEPBYtFCsOiLyBardx9QYli+2eFAW0o9XYlccq7gUOtYtNA6Cq/1IljvydvA2kbURhQ/jdjfztBcFV4W2SJus6OJtqxoD2PYGse19UIO5ZFFC3mkNzb1+N8NaGpfy212HNKGp9+0k8mznptQWlm0kFb6A1OXP97EJoIl3WaHKm2pASJJePIt6yaUaBYtJJreCKWeOq6hs/fT6txMN54pVs0zO+BpwwhTTfmQKlFQkVm0VWT6IOoy2P7osQ5lDI6qQmv0kcWOXdrwBJAAWUBfyXMXhMl6bvxX0vGGOXmaVfosU5UcIjl8uN3feZWPNdwRfedH3pr0ZJz+2f/7H/+L5xIulb7yc8JL3gLDz4uaYnNH2DFFW1pMsTrc7BEWGVgu18QEUkyeFTEIZvuPOTYiq99NYYccbWkhRzXEcrksQKnJcyyGAayOkGQzsxjcF3ZE0pEWkVSELacrA+ScPCOjbf2EPHlXefFrNWwWGrIsNgTZBsBhxyMdyp3B4lyyBlSsJnY9eUC2yVMw0JDX50LOBZ72ZdxhRxsdSvX/vrnjWqqBJJPnWeDBrt/lmI88BIuxw44lOjCW2D98nAsukF3ylAo0/KGPFnIRqztS6LAjhU6XSKEF1JE8AQINP4OIDXIxhCEu6LDjgk6XuKAFNIw8W6HNMROeaPNAjslhTkNw2HE/R1oJFjigUg/OWUDpyDMRkHCn6QQd4hQDhx27c+Se0mxFHc+m34IiiNmmvJQy6LQdrsOdPuCw42+utEoaHdDj2/JbUM4wxeWM48/Z+F172+TIPHnGWz8mn9jjzPs7v2aUvybGEJvLDrG5kmpsUEeUPoitV1soTJjiwoRK9Ppcb3nI077cuuwYmyspo68beFwLLpQVTHFZQSl3/S65XOghWHFddpDNlZSw15E+zjUXqgtmWV3gKX17etXvv//myriPwtX+VMyVs8LtdeRt4/sw2qV1XCN/tfajIoeXWc0bzhsVygQ2FLZ98KM4kQofgrvkU6fbESQXsa3rpHI4JW+qPrhpC7fFU6DWtKBsYLoYebr2N8lw7fYRbXL7JLmqNlkfX7T1mZMvKAuYU4x8Cd0UPB0yVvp3ggI3XFdwsqFKYM4w4vTdRvDCy0EjVWytPqxgqJgbKygamOKXX0jA6uL+8HNs7El35gPXDSqh+wJGqGhQCWRxVqGCioApfktFH1t1lrD5KVKFQfucCuRYVqGCEX9rghGqiydvnnwmSH1NumK/DhNQAs+4eKMZLnhwmxsuGNC3TIxwXT75/vOE6/J7zXBBdYobLhiytyyMcL3a5/WYPzu6bk8t1weYQJp3FTAYm7dsjIC98N/vyEcE4liDZqzUXH1gCeRxV8GCMXcLZcz9gnhMn8vMdW6rPqQEsrOrSMGwu4Uy7P7tNg7Xh4+5FtjmfuyBcVVpsD64YDQruyE75OILht2tctj96tQFXEr1LiAkkq4DaJUBensCCGjVYf5LY/vLR7O0me+gFApp1WcGkOt505r0MFq9V4YU8/uaQaYOLGssmUN4ebzYJB+rqlrtwHC9NUOKoQyBUQmOaGTAWU3qGEzZ14hlpcfYeMKwvzVHimcHfVIJldpVxFlNPpmAEqSexryn2BRCncBaIKWwo6yphEQUOtGMnV9GCchqpLHQW2wiochgdxEZhqOGKvI5hgAw+1xnh6CvAxUFu6wo4JnbusmgajwRBGLljH1QUyRY24Mzcuos9sQGNQjbQopjFwFVCYy4ZM4Z+8yoSKhXPZbHcWRDCXUL20YKZRfhtQcodcujM/aRUpFgcQ9zZT6ObCih5mE7SKHsJNj2QCUCWXXGjihOYURRI5insWSTCaUT20VKprjSy1cwpBOcOPTYOTuYOIXBxDrpQz2UxR5jcwklFxur5CIoFPcAJQI5d84OKM5gwqZeIk/dxcTRhdKLjVV66S4yK2ESixQ8Z8cWZzC2qBfMcp+x6YTKi93lwEXbCGPejrfX85KXQ1WkucsU4oksaley5+yo4gxGFW3jwf/Rjw0v5K325ULtxO5yxkJmGo1hGY/eeUFmJu/Ctf/+y06QcaTVfKaQsSOUMxihFMiicaEO4ujQQVrQVZ/NMMCpDE0WxJwdeJxRAo/v5z/EXKxBpcTpcvaiR9bYQtwAOUMhxs3ZccQZJY7IzRiUPxzxqkmHv97tPePNfr0L/uSRAbr07ndZxxvHPwe8XYbr/dYznntRuDn8bXf8zIm6qyBeBtviU4uPOq00XnQTBisyC4TSdbqX5HUfvPXhP7bLJJBc+CbjV8bF2+uz0XnGvb8OjX28TxcSg5ilOXtqTp5aE2v6LNs+k+5bEjvZeQTMdKzy5x6zW42NR3hehTG9imvWxclDzn1L/nST9M7jtHtuC91TXo2rfaQPYnbMcUaPOdIhZlfc4SSRAV89c2TI6z0YqNo44pWg9NnS/TeyLUg2yfkbauN3wY5QzuinwjHxm/UeG1+o7zjit1RkDSDv96T0dumLPDEunj4Xhvdd6e6XYhXrU+uMCzKpPv+M14Vl1oc3SR8+MR4R4NO+/hLxWrBgx1bn9MPw7WypAUUGfULQNS4MUJ9yxC/iQLvJqtnS8+V6IjSogW+0Fuz48JySe/pbB9VCldtTiStgU1Bbc8QvF0G22epsSZ/chosdLJ7DYDE2jrPeqy3eOYWanNP2Bvi3wZJ8/+8em5Ps2ltKwXb6B3mKyH5PWhrGLaI6fKrctzsylVfma98wk+j/01XgFeJbcd0S0HQ7clRpeQn2ahP7pv5ItiV02UnTdfHVwa4DgTL+lU6phxlKeI64hJe8mr99CLxkwfYiz7j1o4TicySPN25ZvkNyQ56bPo+M1i6twdvx/uW+1ZfmjhEtLD+lc1x8bEJwyEXwtDA3l2ft9GsenaNAAO9MgOGS96ZQ3nPE5T1NjAlXgEeOWPOJP1WELWoJg7EYEcKgxOeKS3y6CGtR6x07ZDy14lVxZk5qQTNZWWB8pEGBzxUX+DSRdrEkz/nxkyKt0CQNpJn1pLHSuvhIgzKfK15pTRdpbz8tyt7qIsyqJww61CKEQfHLFS+1po2w608MsWtdjNn1jLGSr/gYgwqVW1Goqt1S7cRXwdZbA47yK1m2v3zM7mU/YfTWX6U3tZDPRV5JD1DgQYpdbJcQEe/ChKnqC2FxA02nHgaYG1UHw+/h6JZGtDqIXGPHdbXYFOo30rmbjtwNn7upZO6gxuG6krnrMN91jGaowU5fSMJ066ljHWHrcbbjuc9uBhUJ6dB1mOxG6HBAxz/VcUEHlQN3Khm6LjNd56iaIu60hsbqo/ymWJhfzXzHd4PiDGoK0unrMuWN9CGiT2Di46SPojfMJNPXYe7rHudl0Wf8xfi1v/1hvzEi/y5IAyEr8qXJe8VeEBrb0LhPgyjF+Alb4ediVm+It161MjvJVjOKmCAbow6TGDaMRvQkokdRF+aIZjAF0d3mavUdYdEWwbXEJE0l+7ZCAih72aQIDrKh6zLfjdChgI5/u8YFHUWBEM8+IuO4ShYaMhzXYRSlrxb58W30JOnN1pJX6VGBl/xFaICOakxAPjW+zGffp3MFO0ksbVkZx/USq0W5iO89X9ncGdQgpuWspPrpLiQTAIRpQ34cpce6yRDtDh+2+7DE1Ot0yvCMF7+FPd4gP8iv/n1qAxadwarXGSyoM9RU+YYz2rG5rFGqG5xjrzDPJM2gtjA1UfAkcgyV44rsl9mOnoxAmM7+acZ4uE/moGyLnxyF8ONNci4oa19xvnogbVju14yFlQqkvsmnPhZstb9jfQ41gamFApbMqTLi/fr28PF47kLB9ep+6Yga+FbjqXF27zg5ob+6Bmym9di0v0N9DqP6UxsHNtTdeV9LlsZ9dn0YwIJhgN5WrKZDtHMYpJ+KJGGog+nd4ed1cp6Un6gu0XkqUaVX6B8ruz5L2u7u3XUh69g5Rj1fMAw/Fb9/kYEZ02l7EUZbL2Z7aYe/LtNX/010R5yyYNmeMhyXLtLxZbRSA8gwP9E00jEim8WItwThHIbip2U98U2wjMKbIFyHd9kji330PXmV0Hge7SkkZe9ilhi6CuKQbCNIxwfJO1Kc47aHjjvMU1kjbo6NwLFDtusdNsro189TpWFMp5fK2BWGq3mUOI4Lz2GkXS5Z1kjWkMiy5JEFw+lyybJHsoZEli2PLBgzl0uWM5I1JLIceWRRYuJSyXJHsoZEliuPLEp0XCpZ05GsIZE1lUbWghJKn8kkq61veHn423q/9mLj++CBeuelltyFOgjTjoqX4T0Zub9L61sdm/BwaoIGQqHM0qcTCYexHkdKiF4qjm0dyhFH3Dhyep6iOFKC/FJxbOuFjjjixpHTXRXFkaIJSMWxres64ogbR04fVxRHiq4gFce2/u6II24cOR1jURwpYoRUHNs6ySOOuHHk9KZFcaQoGHOZOLb1rMHg3kRhUtucDK1oPXUkdIIGaGAT3hcjcGJpQdEkpLLS1u0dWcHHCkVlkMpKW590ZAUfKxTdQCorbR3GkRVsrFgTihIglZW23tzICj5WKGF6qay0dbVGVvCxQomhL2SyIscPSvs57eanR283dXaHB01tSzTQI1bzq0oPJeQtlR45ntFID056KBFqqfTI8ZVGenDSQwkoS6VHjvc00oOTHkr8Vyo9cvypkR6c9FAiwlLpkeNhjfTgpAfGiGcTmfTI8LmuvffpHYH34X3xboy1Txq2jwaHEE9zNHDU5XY7awLjx3I5kuF9jRyh58iEsWW5HMnww0aO8HME485yOZLhkY0c4ecIxqTlciTDNxs5ws8RjE7L5UiGlzZyhJ8jGKeemTI5auuvPffiXbj1d/4gxNPK2/ZPgSN2K6rsw3bl9tdlp1omjG3LJa6tZzcSh4I4zvN0IsTBeLhc4tr6gCNxKIjjPDInQhyMocslrq23OBKHgjjOU3EixFHi7lKJa+tXjsShII7z4JsIcZQIvVTi2nqgI3EoiOM82yZAnEWJ5VsyiWvrq74if3OMCGhg7pmhsLR/Hbug3RrotYTorRJFiepLJaqtLzoSNVSiKPF9qUS19TU1E/UZkCDb36z2QP3aSNEDpHLX1uMcuUPDHafXKcYdRT+Qyl1bv3PkDg13nL6nGHcUFUEqd229z5E7NNxxeqBi3FG0BIGbnOq4G9YdhLUxDX03ETr1ZXEdWBZX5F6nyvCpuI/QsijagcD1TsoJe52kO9Cmtn4ZK7yGBsrqLxh0ut13KYGyc++wOaMoBi4izl6SSThct+VM6BbDZbgxsq8zfGO/DVbeyhcG8tonzfXWyzC5ZGupjcz6Owwdgasv1ZB5HlY2mRRlYSpK5iX5fRTc7u/yu8lQ3bjJcYW0pBs3s85YVjqjBC6W6zedWT26sBYE55XRlk0RDmb4gOp9Za3FQu8CO69nAZ5RE5/GqgOqaJ21KRrDHB98eJfbWkqRrLr1t8k6AjcHK8WVY/G1KQLGAh+uF297uriYZwE9vkv/4Ln19826UNvXA96xi9jYQf1iPsGH3ddkkvEjkXnSVope6X004GfW4wcDe8wtm1L8it3ERhBKGXMTH4JiET61+GkM9LlWPXow0KcHvcZ4nw2VjLmFDzvqfWtaoNN25ZpbH/VzYdRPD3JZB7GBgxLGXFjCUA8c14XvJa3mq17ltFpG0dzj7kLXQ/wed8uGmsRcWJP4PiR/n40ARW8t/PaEyavI93/0o+SQ5WW49neeUZxw5IRN5IPxkLXEMAEVWkMq7kKCL1AYJTBs9aPFEymBisRcWJFQRZnm+MgzQxGTOAIo04lWOHniIlCUmAuLElxwVpKGhz0JWgA4rZPgVEZAhMVZ1tjO06AD1Yq5sFqhjrRPYyKEXCKZCGHIpF9AOaZCByoacwFFI/9laFxH3o7mPlzub7wSmBehsc49huR3xvaXjxbX5FeTbSckwQqBRmsfohmwW2CE1rjjeLGG6dze2tQ6ByoPcwHlQT5WzJlOFlYKZ7tGCJFMd91iJm1oPI9qPY1QkFgICBLyaaQG4HBOcNoidFOY3qQap6yt9ShBYWEhICzIR4knc0nNHMfOXlI302HJb5rCUBxvApMDNYIFipPP4wWzHdko6Ac2XAv7LOEldsms5UAVQS6SbY9Oj0jiR5KzxpcoklClkItk27PXI5L4keQ8lC2KJJQ05CLZ9lj2iCR+JDnPa4siCYUMuUi2PbE9IokfSc6j3IJIulDxkItk28PcI5L4keQ85S2KJNQ4FrZMJOXcHTveMyyPTxg7FKj75ELxQi4vcm6LHXnBwgtFXpDKi5z7YUdesPBC0RCk8iLnRtiRFyy8UBQDqbzIuQN25AULL5RwvlRe5Nz6OvKChRdKrN2RyYsc/2i8JVglQbAihwhBlNC4VILkeEwjQXgJokSypRIkx4caCUJL0JQSeJZKkByvaiQIL0GUOLFUguT4WSNBeAmiRI6lEiTH8xoJwksQJZbsyiRIhi823gXcF0vwnLcIS5Q4s1SWZHhlI0vDYIkSg5bKkgz/bGRpGCxR4tNSWZLhqY0sDYMlSuxaKksyfLaRpWGwRIliS2VJhvc2sjQMlijx7KlMltr6ceONusIkOLCgUZ+H+QRu1Z1RYuBSqWvr8Y3UoaGO87yeCHWUuLlU6tr6hiN1aKjjPJInQh0l1i6VurZe5EgdGuo4T92JUEeJz0ulrq2/OVKHhjrOg3Ui1FEi+VKpa+uZjtShoY7z7JwIdZSY/0wmdW19WM338j4zXm5jLwgNMn7hNi2ObvgbI9wTUIz7MKlfRH6y8eMN+f+8Q54l1dOPzz182JL3F+dX5/2+RYJh4VWByMiMEv2XSlVbH3WkashUUXQAqVS19UHHO8T7oEG2Hyp0j/iMohtIZa+tJzqyh4o9Tm9UjD2KziCVvbb+6MgeKvY4fVIh9uYUtUEqe2290pE9VOxxeqZi7FE0B+GbeD+zW4wAd3qK1xevWZNx97PSe4zmFJVBDWef0lUxnxlpki6LmVO0BeHrmsc7s0S41HxzQhFQGfeKK701a04RIdTg+fneZfS5AirpNqM5UCysidaLPnDeZoRweYbJlkhuMpoDucKaaL3w41O4yQjxNEe5vRLJLUZzIHFYE4HLxT+HW4wAV/3fXlRECcZTENxgNAdqhTURuDJ8vMGoHYnaby4qgilwU1uVHqA3WJNyhYMXpMHBmu0DkN94AJqv/chbl6h55x0+kAavkhVovd+mh0QeCn9+6te3b6/bbvqTD5H+/enBX5d4Mf5i/Nrf/rDfGJF/F8SJAL8iX+cbfqbQb4+CfJCwFtyS0VyFT5i7++SVjbtjC8tUHN+eDcLsycSCIFxF4Wp/Oi6TrSAJDx4XD7PzTVYCu/oKCgsQ/rcmrk4UunmALBhelqYM8tWVvYpx7P5TLgZ50Xt/HRr7eJ9Ez9M0DnP21Jw8JQv29JmR9vbxeflkYmzIhE02V60oEts9KQZK6Gr6KiFFKk4gNI0/h2O4AHKBNZlys/pq7d9tOsB6W/z7gUxcpzajnblMgXvmqziAqL41menF4dOdvJpIQjV7mUKXylcpKZFRO3+VGOCZwIA2YE3m3MTmY/xNsGJs2iG4b33yluSV4+zlCV2eQZ5CuiV5fY+Q8qO/9U573dPZXi+OPeNlfP/EsCa/fLyUDTSVsGrzjK+poGXv5sfkKw8fkrnP8JKX/M9LzcgJXRNfHWba0J6aLziItMFjEwnkAGuyQEjkL//57a+v3mDiENeEB++Q10TfeaDYzMH4vjlBydzXr0fiWMTB6IM24rJhYvMGg/+miZC3y3C72u+ChzQ6jwk8+GL6sBPKAFGJHegUNn8w5G9aCPnLgpuYwCu8kT7ihBRzlcSde4ONGpQFTBshalTRSRtmjfqSasQ6RdtkIpb1BBsvqBuYZd3gnb8x7qL9fXgORbAw++E0Kn56UC71re/9rR9EMCZzUQh2RSUCX1L/WMCTrVGlOLKLrtJI3DaBoBiOK57pI//tpVE83wjvycCcRo7Fa1PPDMEftlrH9ewJlCRMFzVlTTu3gSOGaP9nCUT2qlhB9cCcosaKuSEbOE8YtnWWQLyuChLUHcwZapBEc3yGAJHuTZslEHKrAgRlAHOOGqDLMPr0CDo2Sh9C8JIJboRg3N5coEboItmhkgG/DeLDh6VgNvUQcKI1UB9aAnGyKlowPG9N2qJ1sSVdCAMJbIiust41TCM/TiBbk+wNnnPTB+GvtU4hsycwvm6Z+oHpJhnqxgSV2GMJBKiqcMDgt2VhgKO9tjeiUUDD7hDngcFqy9aPhhQZTjcjiEI2doeQDYw3W63jzfIIad7KDoIRVBtWu308xoTxYks8Xpxl/RkvLKV7D0Q+TqnBiBcYtz0YMOJrtY74vtxH4b13VhA5Ebm49yJ//UPSKX6axl2U7ga2Va12AV5oaMfcf8sJDYzuWq2juyqhCeLwPnm67HoJvZNTbYg+aihuDi81MKRrtQ7pqqRm+/jTgKbSDm3MUE6aczMDY7hW6xiu4pnmZi+/MIuOiabYDn3MUKQjXmZgcNZuHZxVO898EsiUm6GPGEq2NS8xMDprt47OKp5lvM3hw2nzOFhmqg3RRw3FdeKlBoZt7dZhW5XUvAvJFx5+Ts44kk8dfiIOSBjEpxql1CtNBsQSX/P0EUZRsHkJg9Ffu3X0VyVhb/xPYB0rNkIfLRRRmpcWGAm2W0eCVdLS7RAFDlgw5OdN20dwLBgPtlvnDytl5RiAP3yIk2P8D0G89wY+0dS0SR9L7eM6Fgwh261DyKQVsX+3365EsqquT6c1/OpTkk7KT4Pk5d3lHHvgKMalgqty9wwhaWZGKTP4W64SXLYFw8x26zBzf2C1Ebbw4YRKq5jBaBA3RDDqbLeOOvcJETX5ZkSoNUKUUoC8CMEgtC1QD/zw17u9Z7zZr3fBnzzS/5feff7JytiXeLqOwmRsshrR3jqsqOpBvAy2xafSMbr0opswzdANBbx8oUqTcfJ+528hLEXE574Pt8nhSePi7XVSjCsO4p2/8aggZf1DPr08d0xyAWDStMdp27LN4emClnIWT7WF+goMdljpYMjaESjnrZMwkWPKHFxdB/fpOCyzsuir4MGP7vw0t9zPa3D5CmnSfWJmTqmky2KIUkmXGwTq2DcNeXPNXduCkXRHoIi4TpBF69JJLzWuGf3TjkMf/DC/DQ/8fLXLbQuKAk5ZFPiOPDQd8ZhO/6tgnZ2jSQsRkJfYBGTc3/qEweDUy7t0kYWWcNpqHn5KnnWbP6tiECkk3oZAHXl3ZPf76ILsBe9C47V388RI3yL+snGOL3VDnWkIlaOo9AH538vrC/Fpek4pKu3nTTO+3a72pOmBF8SkZ7/37sOIn1fyZOu4pYS4squel4e91VDXjS7/oPLM4lB3cOyhQ3y9j25qzorrQrn0WpqBhvHkIQBd7MF6rKFA4jhDx5qnroYutkVOWPQCuNDlO2gAB91YS7kNpR3HHTzlZJfGOg/die5y3ej0a4xf5f98HTxEPne96LTe+FfH3btfVz86jOkFpDmt6tQXmq2JUpd8EOZ0HtxmY4LaljMdujG9SxOvWLfPMMxJ6MLobUiasFyTT0fe01sveJ+awLk2f7MX22QExTboNgPottbdJl0lDIpczmzohF2sb73HW+929wPzAjcuzHq4/oEPN9Ag3cxB31OEOaiJOXMh5lLKtkl2Q/CjR1HDLsmfJvdO9Rx4MK2uW9pVaj5kslr7AeFnQ+AKjUfLtDmxcfHV868uv0QajqCU67LeWyLLMWVUiwMpNcZgQ1FNDYPPRwY/DQafy2eQIrspYfByZPDTYPBSPoMUxWyhgsGua3GHQJMmHNGFnyj1EHWt0IIxJYqqpQTSrov1CClqSAWXcEFIKaqVEki7ruYjpKghFVzjBSGFGpQ7UQFp1+W+VZCwzV2N+RcZvzr92/M0YNh73L2VZSCKQ1pC100p3V2URrLRJhyoWKmxia67i9EmRptovZkRtAkoPKmxia6bmdEmRptovXcStAkolbmmkE2k5zGTd7tYB+fk0TwPdJXUKYyNR5fkZcKt9yWwmLNYVhTXSp/vM8bX1RO4O3aGV+2M5bkzloXGYYz1wfIddToYHWdhKjhAkBojdKBe54olFWMGv4s7rNcG8LnFMINzaOYg5k47UEV0xTKVMVvG4NIoupkTtiQLSkl/gSQLB4qLrli2MWY0X+wj/0fP2IU7ajUOvHM2eHHNkFEuBxjalJ136fWxS+tnbKh3umLpyZjN4pUfieYpD3KqPrdTt/nAA7MiczTUNV2x9F7MMB696++S3Mazjz2Yqfr4/vvk/W9Or6+bOLF0X4wTNgOM+nkbqquuWJ5yU0DydRpOe5OG07a/fDRL9L98cv3kwih8pEvopPPBJ1oIEGmExKaUzhE7pEEZN9pQUUeIJ7hBUUSVgmWNYA0FLKsLWC5FVpSb3dl+xtJ5apOLMXShLZtSv6m/eUwsKuVSxDul5AlMaSN5GMljT3SC5FEkMrk5cO3nvM7uSD/8lV4Tj9cBha4ep782DoNLEa2UoigwCY4oIkeRPR+2QhGqRFOhZJqY4qqTHxZ8ddIkjzRoldw3AOM859/mSStX5GNxXrfqXCh0f7NfnyMBsXf4sEpzU1anP+/inHSO5hRbcYXVK4G3LQhHZ1qMdu0Ac44rj0sDNaWpUAYMSpB1JgDUM41vLyomSWLCW3AjC2WiqVDKC0rSNaqm9aBjU0UdCaqoJs6F1FAXClBTofwVlJjrF57qaS8JSwaajbUjpmVigr7VrhwqSlOhBBmU7A8nA6DeRtAo/E6n8kouVJemQukmEiH77vpliaxcsL04C7b5JzB4ePmrIHTxHCg89TxN5n3DHD8OT24KxampUOaJTi6ROGw0RNF5bA5UD5DRKuaYTaG2NRWS6/WCG5FhXhu/eWdaC0sFuZV6oqXvSwqL5j/o/dBao9mAjtFtNxJO9qi2m7TPmi0GanJTITVYp8VIqBXa9nhnwZdAZy+IjnM62iPeTYbSxjmcQvVwKqQe6jSaQfqANMzxOIGd6p1Oof43E9f/yugsveguWeDJcH7Y7im53cePG5fPLw27ASAyr8HPa3MDE0M5J2gnjX1ybqnxiLy6aZAxwXoo2JUVPG4ecThonGPL4yJCsW8mLvZpo1aj1tGaX2wyiCsrIqyQZCG9YwplvZm4rKeN6Ys1+bsgu7DNIKMeDONYbmt7YLRXt1V0ighPoeQ2E5fccCD4Q7CKDj+/D1ZKHDSkHFYarRtGsWh0FUaogc3ENTBtMOoMBLeGCV2I2JUVIla4ygsGiaHqNhNX3bRhrTHodbHe6Shp1taWEMXBXFkBY4VWVBzlRiOaQYlwJi4R6jOi/fo22A0jDNae/0IjdfMvFgeuwgZlvZm4rKcNtldhfOupgU0FNsXX1Y2NWGRVy7RZ6K76GRMqbTNxpU0bxO1uKNZHMaari6cDiJRxX1Q8g+LXTFz8asvxN1fPs4tyOTk+fV6jXnECGPKbvh5WqWIqoZiMKLan8eIcVg6pYgYFtnl/AltnYPWGMcTZRRfBmMpKyVSIsVgEYwbFt3l/4ltnojVGMHRVZW9hR4iiF1MNMUBRCxKr1j6DUt+8P6mvswXp8ijFKUbkTE41xOBEKeZ2JqFOOO9PJ+zO72By0FoxjyUxbdot6gblv3l/8l/St+ZEELLsD/C6euYEra8noTh/m8kxGzHekeXx9qC6N+9P3ZNALWp/jwYwPodPVh0KpSyLuXxzqLfN+9PbJGA9On1cxoTI65tpCv+JmZGY3zeHSuK8PyVRghkh9fxoKCNy/WaaQoBiKPM6f3OoJM77UxJlQPzpuH9U8LH4fzOxuF0VMyj0zcVqZJLlkjRwFyz3a9LPxqOX1xevv4ZopT8uAVUpa/LuVNYku2mgi5PXvThmuVHEkNZelKNknJqC0MWbwWiAYGlM9nie2s0xchxO3Jwi2Ymhdx4UT4C74+uS7US4ufFGzvRwxhi8U6Mpw8QDFVTNFmK3Fvcxn707/LxOPAv8tJXeVDdzMFFMy9x27BOjmUaoQAnSKH2KG9HTg16L6U6AM6gULXhzBV7t10F6VVMLul6SbXeya04vkypS9t2m39sdjo2gk1R8G90gwRA8N0iUkaLgA8ak0Pp6iKAUtCjL5e/Il99F+/vwHGcLGYvoqUzhC3+5Dk81CktkXRQ8tKhEWFp3vcjTJXlGaKyDh6gTU1f+NvlGvwTUVeosbpMJv+gxPjs1MjISq00dTd8I70m/pstD3WzHaH45pF5pk24yYUCdSWaVHCjHLASF6kfXYRTlt8vFt+SXCSQh5b5C2seMCzIgH6sQ0T/66vDTTZQHX9VqiMZfkjhClaBVclA9ebHrr50UsZhAlwASG969vyVdkQYv4hKLjUendskjUyqipLFPjOR5G2+7zKg4faPx98Y7f5lWWyGf9QrdhtDHmEupolFDVjNMzQxx+CQLKOv0ax5vTiB8TzYhWb0hHtflCE7leaPpjKYjBpo4XzxmBWWehWDqitZVJ1n//Z2oJ/bJG1CpW3SbkZTKdopXoGKH1RsMlJT6NRiF69BoTKMxqV6T+A0NimoLwfQ0vSuT6uy0IVsUtkS4uZTj06rXKKE8uAVUBvu1H5UL1Whbo231t2SJ2R1FPBXMP9W6bh3TBL28uM1oWuDCqVPX6LYsKUcRFK9aQsWSFhStt1fjUbhojYY1GlZfS5aY0VGEb8FUb60rlsqCOwM2Kkx1feZSDhopXqm4C/ssKCJ/rwajcJUajWk0JtWrE7+hUXIiBM9vaF2ZfhPtwtvsrIpBcNlFea7PaFh/pHXRo2MffYnBzBZSarkoXrNYhNWZlTOh5FL0alYK16/R5EaT07eytTNHSg6GYJ1INKtcmgE62l2z3ZU6CofxSan61ON6B1irtzJK4kavVtbTojda4GiBGJY/QesE2R72RLD0rPY8xFZnKz55m0N2QmMxlGzE0Hh97LN6ywF5Hj1bjuKExNGqRqvqJy2R2+JAhoc9ESwq3c3iqjdZM5rr0ZtbveZaR/HH3fHFIsqLITwKstAjCOcd0m58m09/OBOQbmFPBIs7YyL5M7zwTcCSEBXMMydS7nnp155EciqcCcipsCe9HleUa1na7jQS4BuTRGtOpJSE7BdwXlnWmYD8B3vS6xkouXDLloyoBRJeZnX1CAxk8o/S8i0bI9wnDkdWaI/8ZOPHm2Tiz3qrWDvh4fBhu9yva8sliNgKXgnInHQpxe5MQMaAPen11IM6NKUFdofGJ75ArTkRO9NQgdQE+rs96TXLWS6kKo+NygIK0xlOczJAb5L33KZjAjXbnvSaHCkZbr2XDAgQju3clznRk6fYDXORs16OCTRle9JrIpXs+EleKy/ex/f+NsY9pTPeVjf15iBDG3lXvit2ZT36FMFWsOZ18i73YRxkb/vy+sJ49Dza77ynF4d/3698aANpUcfTmBnph0v45+XVvXN59UcX211wFxqvvZsn+VB/WUS+/+B3Qva51Y+ythpPjTfhXYD1ziPTlKDv8A04a4z5h5Yj7m1CzdQULHCMAV6texMxjvHtTkwJx9L7QlpwVwL1SVNIn8RCd6SCafk3X5QtIe2sRmuIUNiAJXbbS5UzqB6aQuohDs4Un3hTClsTaKg0EavLJSuOCRU1k1dRS1+6cI8PwOry+SVpv2Ua/3x3+3S5sWhIXWZybbJNTp+XfY2GWy+y1qyCJB5TeOe0fU+MpCmPXm530ZO0evUqjNDuKq0OBeFpQ0obxdrB49kqQq3L5NW6esCuwxZQK4H49oNWh6LyUmAU3ORBocvkFbp64PKYLPFdotqcrxzsxUHpTufx7ffJ29+cXl47ot0CqlIoZQxsLasW1LtMXr2Lk1XzK3dAa3fhbYtUZtE6tMu1DethS5ghKyPXdb22oPxk8spPfbCmc8EWxA7fGm3D2HrPBIot0hbUh0xefagPGPWv0oJIIl2Y7W7ZfnLQbLUyQxXHFDwUuypdkJs15R0h4pSsQ78cN2mfTSOWcTfu6fMalJusUXHeKOMReUvT8AwH70otFtfjusWZMrJwcDjHkGcdhxKNJXisTReZemQZfkgRrutS6n6o41Vw1Yf6iyV4PkwXuvo3BDwAY90FSDm+r47iVhsEqPFY4vcZVoc2uQ9+ATlm32WYtzA/MXXWd4qPkrMxUH+7IbszUG4lxHJVq/hA0cayceMjeQpEwxPWSVMsplnlC6ozloObr1d+REZ1F+6yi2UHz1W1PbqBcsRu7asCBWUVy8UNlDR3Aw1S+BwURyyptwKVDfUPS/C8T0uHIOngLLOR0yM4/wGCOEv2Mmi3Ro6UM778+/3z2PCOIUecxYZ6iSV4XEcfm5jcVRasSPddTs8hQnF427irNhRcLMEDOfpolrQz7CEXkZd9dHvDThmINhRLrH7EEglw4QhKs0BBuOPrOZgnPj+KBaVtqKfYvHrKW38XRmR6SL4xmSySkbyKDn97/PIhadEph6oEb/IB4/wBGrzFJO00O/s0eutuET2rPbGFNt4nTfRJC54Yj+6IX7P+MskhJ83ZHX5OqjqdqEO7QYWntqz3Fl+6BO+YU4aZc3R5NqdQSbF5lRTd0HaZcHXzi3A6hmFCRCgLTsVQU7F5z83opvpifes93nq3ux9C0U0qH9QSd6lSTAE0WLcluDC+ybSEKnhQjbF5j9DoBq+zq8+XPYkNP6TRA1cgEbMKIZRsbN4DNUX4rn/wo024Nv6eIPmjn5b+pXhQx18lJZazP2iCsOG0YPsJUCeAlM46kjiAGQ/GWblhg3KOzXtKBgFsPR5MaAkLEk/HhdFM7sRwnnFmDK3UohQOVIlsXpUIAat6Dja0xBafg+PC0CQugsV8HAfKSjavrIQA5kHvNkUMAev+EkZfeZd8B0pAdlkC4sngOHdnilXmwhD6SXPzmlmcmRzNDxpK3mJzSzDuDKadSpw4UPKxF0OAqZea1jrxQlzBetqpgrUDhRpnMjTk1NaqRsMdwsrU006VqR0ouDjmEOA77mPicH34mG+zpDJ3vihsG5IXX65TFffpbbJlSmqjn3dsMdnI+RvvmaETU2p/6GZz1imL0oGyiSN+FEUDmwNK0BWkDJ8bPRPLyawiBgUSR/y4ipbpT2blZnScIS31POvmW0AhxOn1Fp2VV7p0lFVE/Kj8HYtIy/NZhSJ99DY8ekHeJFgntz8+PTYGbZR6pud+58o4cwwvT7QaKitK6Y1pzToP/gjwpwMwx1DLYdiFiovT62VREmZgmY6P2OWXx2/UeAtmF5PC6SHpuTJFfIUQugzThWKQUjtTtFaMpjaamp61TMzaoADm9Hq7nIRVTc9plS6M4wuFzPVcQSu+lohlFrhQklOKt6LFZCR8KITLmcIFIaeIgL1eoijFM9F2m1y3zQrKWONcynntXjwD4bvkXIrqqJJ2Zf7BCPzQgJe1PxdnnqJm9np1qIQZXuXNSdIxR3UV0nwoSkCx1+p5pkinKnlWNIePSONGWs58zU81RaMVLPmhfZZObm4PjdfBQyQ/+eRllvVEBjzcpklwSeQx3Cc5AFniOvnJxo83SXQx65livsDD4cN2uV/TUwQ6uaqVJmu3jU4ZoS5FalUJoaqIhxwO+4t3DBsivdGOct/VzrFTqMK6gpX+9c+xXnQTBqvI26lx+sqyUPHbjF8ZF2+vh6IEgX7SblRi5a31bcTJIDdbEtRZlVqSqoViNKbRmFSvUFz2BJVUV/Aij0dvwm1Avp1lR5dBtNwHu6wEyE/HEiAnq8kufX6Z7q295P1fe3ehJ2AsiQ2E8a7d8VRMhsLuibY2sTjZhHc2iZe3633i5SSWcJGO4MXmJiA/IcbBYw7pQx9ZX+YmYdnPjOqvT78CxsJ94JsNF50nduflZnDCBJgAVFtdoTveezABag5nJ/C7ENiYqtkLdw6buw6FBuRxl/USGzuof7q8lar6wu7FP/4GE3fH19EMnssGT+BMvzrw8m5ikwe1SFcorz3ZmgSrfVbxuLxn9rak6eVdcfqj016s9LfHgX1JNijUWlMM1KYt1vZrP4wMb70MEyFrGZLlPfvW/lf4qNAD5WB3+RURsG7PmKzb8BSu894RYL1EwhmcOl7OnLDhhqKjK5SIKx3udtvXNohj2L4y4Ua2bbXnbLKhj6ec7MZNKtQeXaGUQOlYi2xJ+WBuQxWKrahjMlmi3KqjnKWmjScU/NxyKtIrP0rgOB9uLZ1rjfxN+sWvffLy+8hjAGVc7deb/O9PYB23K/AjxwFNn8qqp3H4N/JWcourpa05vdY/GiZArPBGTMxMiwx5Jfh0jn8We4s77JQ+MqeovAqnvyngBxfimk0nGFvqeMLxYY/cuX/qI01Q4nPnaLCjbb2Q8Ce0K1QM4rwWRLhu9gpi2lEXpY6qJXJGUeUWaIhkHtvVAmHTOquYu0Utd/CYU6/ccRzDnUHZajpBg9pVRPA4fNiyCkByEdd38gwHtOV26WLXnNTCa8Ji97x1IGdQvJmaCqja/vLRLCH1zjt8WGXXd2bPCMkAmHK3cQ0lqh78KE60r4fgLvnUyUmtYSV/6/xNse3xTIF684Kz2HH82MPGuYebQalkaqkBzqLOYeVXt7ATd56dLJTICVSdF0fOKq6alIHjZQ7qJFO7D+ao1oIeudMkh5M4IUW4FXHsYeMFDsojU0cmcJfkn9Hhw+3+zjNckw6dZxQ+JX2v1sI7OL9Y/soYd1pO/U5LSBbmg48ylswhLHVSPYJQxJi6aBDUEqKr5w/D1FYfpDMVROlE4OOd/KDOMJ2iIU9vlK4eQUxhOrM+TmcqCNSJsCgcqIOKxVSqYtEJyt4DdfUcao7UmfWhOlNBrE4EPZ5YHRQqplKFik60vfHi2DNif4lrA1h5LV30WfWxNksg1qaCvqyX3uW9VEvhHIoTU6niRCcKf5NfdSmbP5m3BXFQe2qGNl7Nel7b30E5h4rDTKri0Imfy8Pf1vu1FxvfBw8e6wSMppksWJJRXob3gWf8XRpGOL7rw+ldteFi1eOiIJ4nMr3BUa2f4qB8MWtxE0f61vHpnfzt6vDznb+FxLKr0OctSi8QK2L6PAred4GzZXhP9r0adT1UYv/UXm2AC8QHqzRBbWLW4u4MhTSpk1k14YIj6mfVR/26IAWlh1mLuzIUIqVIgdDEE4IoXuVwB4BJIIRchQnKCjMHFUwSo2uYqMIUmLPqA3Mi54eqeEHJYOaiwqtbnAwTUZpDbFZ9iM0SiO5WIYLR/9kUFURywl+YYEISObMbImcCcdsqVDB6P5uhgqprNAsTTtpDWnZ9SMtun+44h4H5WYtrrxWC9MJPrqy4T9yiYAuf/rmEEZq6QRuaAvGxCnsLGI6fyQ3HJ9cFkj/YlEC72IRxmjcdnTo+NEofbfQEuU7ry89DO7/usvi62Pw9W0FCWml8mkfw3AvMI3sLGMyfyw3md6RP9AC/EH1+7Tn+p8dz/L6x36aHJ58ZysnF5FPaKtLaxBDmOJe/gMH+udSzCl0Rbl9iR8/kqdkJtYVKnSihruGc8wIKAnOphxW6ElfryyImD4vHCovF9k1gOZ+DTSLUEeZSjzB0JZHhAHNVc9DDoH43FwZEmIUdqjhAJWAu9YBBfiLiYh2cn0JtOc2pTc5LnMpMtI+Q8JX8kJ7vc/2PSfu8UsuTYfdunpRajQsmpz74RpnpuFmDssBc6kkC+axpOV7QihsE/qvTEG1TdzSeb8RpA815FGEBxYi5uBjxJv93I4zuyPckHuOWLNn7ez+6PXxYBj4gmh32+5YsYNE2rcznGzdh1lp/++DF5Uuy/tW7MS6ur4y4/C38OHNcSyE1tMfVSSX0mW3UZQiuXWsILsykzOo2h1yzKJQw5uIShkwUX74fUUSLYn1Cksu60pMPRSiCzFWUUZJZ80GlL/vMuA7u09Koee3CVUA+epcYwy7y88KFxKcgm0t/41GRQ10donh2Bn+FCHcCdZK5ippKUitEDILPl0hrSRQBHUA9CXcCtZSFilJMMutJDAJQpJUninzirz7hTqBOskClk2AVmr1ztBDuEZEAaCtYwaWrze4ECicLVMLJp6g21+KLQm4ucqzioJlsydmdQNllgUp2wSg513KoS3Muoqc/Yaded3YnUN5ZyK0f9QnrzrX46RaeixjqT7rhE5/dCVSAFnJrSX2S4nMtiPoEwyKCMAOHSxV8fXkxMatOMPnh46v8xycwXns3YeTtDj8nF4aQV770sgTeUq95lMzpd6TrPOM3N+vDh3/ZE7+L7JbKvvCltyFtOPxfPzYeJU8wrpKQcXYB5LGTv9uk+ysKGp1l58tTjJ5yK8rieCvKP6QNNu4jnzzES4lI7ma8T7gwvPSXz45O8zKn9bh9THE7/G8jKPh6z4xpcpHFfr0jXxwn32U5PxjJs41UNvjTPrvpMTbipP/irN8YanfpzQCgxb6jM3p5AS7GSN8k/xLy7GCTtZGJ5tZ/HxrfvXtjPLr3gti78zc1kcqaWy9YmB3/yaKpDqJCB9BnxtQKqr52bgUWIit4FdxEyf0nHs3f6c0Qzm9h/Cr/j1VGx1enyt1Va/lVbxaSv9G1tw024S5pCmnS1l+S2Tx/mVZGVOl6HXYE/aye7ag01g3WVA0b5NZkt7Km8rhQ7CrM28GyqCQElz+7PKC7HBMxe+LYAL8q3JlVuHcroXhyNqVkm/iVYRv3ZGe0CjwCr5eclikIvWRtIf/vkS+ODh/iXb5fqDCeXMeV3RT4LJNek2/cb1dBlFhawWA7oF/qKR34tyzdkNBoVZ3/jEbLaUHjehUWTaAE41WUpKoZ3wQrMr3u1/lt5ycgWV/QvsKM3Ik9+X4ymWZ5BOR3pcYaTw3YvGdtZnJC2kMQ705z+hPjbZ7vFx97hXxZZtgpzaT/iYVtiCn4yTeFZNt8lzkoLJ7z9z7avEhdG6UQU0pGtJrDKxDSuWP9+bED2LO3VY1a5Pbi4rKXK2rctzdzuUrvCy16AeQnX/EbTroTOg6C4S39YEd+XnPZInPnUof7sZN00C7glCqj/aouMJzCXo2J5LBPW8CuauN/8VYn6BdvK5xfvM0xL8zU5Rb2x3feNTrwFsjEVrMhz9peg3Y1PyxHe9ECbYXz+MWSbHh/NMgG4fbwMQrah6I7g159kSr36e/RTfH0/tNhEUIZ5oom/Kw36g3DrmamZYZhT3AZRnMRLoZZmFLNoqp7l40iE1/RGQW/TK7UJCglTuoSfJVZRZNEnloFXQiw2wgBCq3ihf9+F3nUKiq9mMPx+yt2cPwxOkso9ZcOExA7bqHMBIr9UGMEdB3AbqMDKDSC9ukiXfG//6YC/v036JCvzU1RC7tQNpQi0uuSUlLG6dF5u010XmV8x/qNq2uST767Gt0hP0JH+qmPdLAOQ/JaJvZjH9QATxcAbGQCwMUmEMmmlkt8+uVVhzf5mSjzxZR2dT7vqat0gA/D+FrAP3UC+xRBQv9k/vgFfWs/mbcwADL+jw7/c7sKlmku2FUqHIZfAmN44cfBduffZYlf4AxhkzHAr6BYRqnhbXY9F7u9tw5+pFpHUak1U412t4+22fGYRMz99bdvjcurkn77JiU6PN1Wnold+dHcB/JNKy81so0RJpJY8rLBQ2jArnqaHCIx/EoG+vlgzlF3e1oU2lrJu4V+vi/0swazgrWKi73+yM7+JX5KevpLYUOroZYOKutRdC7Za09ifXSfol/rs0brG61vCNZndbO+f/r/I/AzjQ=="
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
        st.caption("14 itens por ensaio/ponto · 49 resultados previstos em 24 horas. Índice de preparo: 1 por turno de cada CP, considerando 3 turnos/dia. Somente vínculos confirmados entram no percentual.")
        st.dataframe(pd.DataFrame([{"Material do plano": r["plan_material"],
            "Ensaio": r["plan_test"], "Equipamento": r["equipment"], "Condição": r["status"], "Frequência": r["frequency"],
            "Previsto em 24h": r["expected"], "Coleta": r["sample_type"],
            "Validação necessária": r["notes"]} for r in cana]), hide_index=True, width="stretch")
        st.caption("Impureza mineral e Bolo úmido excluídos conforme orientação do usuário. Fora do total diário fixo: análises por carga amostrada, por frente de colheita e uma vez na safra. Comparações NIR não são automaticamente resultados da rotina.")
    content = uploaded.getvalue() if uploaded is not None else monitor_bytes
    if content is None:
        st.info("Importe o monitor de digitação para visualizar os cartões e a aderência por grupo.")
        with st.expander("Critérios da avaliação"):
            st.write("Frequências diárias: 2h = 12; 4h = 6; 8h = 3; diária = 1, além de 2 ou 3 vezes ao dia.")
            st.write("A base contém 313 combinações comparáveis e 207 pendentes. O indicador é parcial.")
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
