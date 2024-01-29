# there are 2 sheets in a file which is for vn and foriegn terriorities
# the purpose of this code is that labeling one of three types of sentiment (negative/neutral/positive) for each comment


import pandas as pd
import numpy as np
import re

#df_bu_vn = pd.read_excel("C:/Users/VivoBook/Downloads/Classify BU and Translate Comments.xlsx", sheet_name = 'Vietnamese BU comments')
df_bu_over_vn = pd.read_excel("C:/Users/VivoBook/Downloads/Classify BU and Translate Comments.xlsx", sheet_name = 'Foreign BU comments')

#df_temp = df_bu_vn[['Quest', 'Translation']]
df_temp = df_bu_over_vn[['Quest', 'Translation']]


# replace non-alphanumeric character
#df_temp = df_temp.applymap(lambda x: re.sub(r'\W+', ' ', str(x)))
df_temp = df_temp.map(lambda x: re.sub(r'\W+', ' ', str(x))) #chưa là df
# lower text and trim text
#df_temp = df_temp.applymap(lambda x: x.lower().strip() if isinstance(x, str) else x)
df_temp = df_temp.map(lambda x: x.lower().strip() if isinstance(x, str) else x) #chưa là df
df_vn = pd.DataFrame(df_temp)


#----------------------------------------------------------------------------------------------------------------
# homogenerate comment
moi_truong_lam_viec_1 = []
moi_truong_lam_viec_2 = []
moi_truong_lam_viec_3 = []
moi_truong_lam_viec_1_reply = []
moi_truong_lam_viec_2_reply = []
moi_truong_lam_viec_3_reply = []

danh_tieng_cong_ty_1 = []
danh_tieng_cong_ty_2 = []
danh_tieng_cong_ty_3 = []
danh_tieng_cong_ty_4 = []
danh_tieng_cong_ty_1_reply = []
danh_tieng_cong_ty_2_reply = []
danh_tieng_cong_ty_3_reply = []
danh_tieng_cong_ty_4_reply = []

phat_trien_ben_vung = []
phat_trien_ben_vung_reply = []

lanh_dao_quan_ly_1 = []
lanh_dao_quan_ly_2 = []
lanh_dao_quan_ly_3 = []
lanh_dao_quan_ly_4 = []
lanh_dao_quan_ly_1_reply = []
lanh_dao_quan_ly_2_reply = []
lanh_dao_quan_ly_3_reply = []
lanh_dao_quan_ly_4_reply = []

da_dang_cong_bang_hoa_nhap = []
da_dang_cong_bang_hoa_nhap_reply = []

luong_phuc_loi_1 = []
luong_phuc_loi_2 = []
luong_phuc_loi_3 = []
luong_phuc_loi_1_reply = []
luong_phuc_loi_2_reply = []
luong_phuc_loi_3_reply = []

cong_viec_suc_khoe_toan_dien_1 = []
cong_viec_suc_khoe_toan_dien_2 = []
cong_viec_suc_khoe_toan_dien_1_reply = []
cong_viec_suc_khoe_toan_dien_2_reply = []

co_hoi_dao_tao_phat_trien_1 = []
co_hoi_dao_tao_phat_trien_2 = []
co_hoi_dao_tao_phat_trien_1_reply = []
co_hoi_dao_tao_phat_trien_2_reply = []

muc_do_gan_bo_1 = []
muc_do_gan_bo_2 = []
muc_do_gan_bo_3 = []
muc_do_gan_bo_1_reply = []
muc_do_gan_bo_2_reply = []
muc_do_gan_bo_3_reply = []

cong_viec_co_cau_quy_trinh_1 = []
cong_viec_co_cau_quy_trinh_2 = []
cong_viec_co_cau_quy_trinh_1_reply = []
cong_viec_co_cau_quy_trinh_2_reply = []


# def for analyst sentiment in each homogenerate comment
def ana_sent(x):
    if any(substring in str(x) for substring in ['tệ', 'cần', 'nên', 'nếu', 'yếu', 'chưa', 'nhưng', 'giảm', 'hãy', 'chậm', 'tạo', 'thấp', 'tăng', 'mất', 'chê', 'bận', 'ít', 'do', 'dài', 'quá', 'stress',
                                                   'chập chờn',
                                                   'cải thiện', 'tốt hơn', 'bắt', 'bắt buộc', 'hơn nữa', 'thêm nhiều', 'có nhiều', 'khá nhiều', 'giảm bớt', 'giảm', 'mặc dù', 'bổ sung', 'có thêm', 'tinh gọn',
                                                   'đang không', 'không được', 'có vấn đề',
                                                   'hoàn chỉnh', 'đảm bảo', 'để nâng cao', 'để phát triển',
                                                   'giá cả',
                                                   'tăng cường', 'tăng chất lượng', 'tăng',
                                                   'quan tâm', 'đầu tư', 'đâu tư', 'tạo điều kiện', 'cải tiến', 'bảo trì', 'sửa chữa',
                                                   'mong muốn', 'muốn',
                                                   'lắng nghe', 'luôn biết lắng nghe', 'luôn lắng nghe', 'biết lắng nghe',
                                                   'tôn trọng',
                                                   'có chính sách',
                                                   'công bằng',
                                                   'lương', 'thưởng', 'phí', 'thêm phúc lợi'
                                                   'góp ý', 'xin góp ý',
                                                   'ảnh hưởng',
                                                   'bụi',
                                                   'sự cố', 'cách biệt',
                                                   'không thấy']):
        return -1
    elif any(substring in str(x) for substring in ['tốt nhất', 'rất tốt', 'rat tot', 'rất tự hào', 'rất hài lòng', 'luôn tự hào',
                                                 'nhận được',
                                                 'đang phát triển tốt', 'đang phát triển', 'khá thành công', 'đang hiệu quả', 'và hiệu quả', 'rất hiệu quả',
                                                 'cùng greenfeed', 'gia đình', 'ngôi nhà',
                                                 'chúc',
                                                 'đi đúng', 'có phù hợp', 'rất phù hợp', 'hoàn toàn phù hợp',
                                                 'tôi đồng ý', 'có đồng ý']):
        return 1
    else:
        return 0


# execute filter context
for col in df_vn.columns:
    for row in df_vn[col]:
        if any(substring in str(row) for substring in ['mạng','internet','wifi','cơ sở hạ tầng cntt','hệ thống']):
            moi_truong_lam_viec_1.append(row)
        if any(substring in str(row) for substring in ['trang thiết bị','dụng cụ','laptop']):
            moi_truong_lam_viec_2.append(row)
        if any(substring in str(row) for substring in ['bàn làm việc','không gian tiện ích','phòng họp','cơ sở vật chất','văn phòng phẩm']):
            moi_truong_lam_viec_3.append(row)

        if any(substring in str(row) for substring in ['sản phẩm','thương hiệu','sản phẩm food','thực phẩm','con giống']):
            danh_tieng_cong_ty_1.append(row)
        if any(substring in str(row) for substring in ['chất lượng']):
            danh_tieng_cong_ty_2.append(row)
        if any(substring in str(row) for substring in ['khách hàng','đối tác']):
            danh_tieng_cong_ty_3.append(row)
        if any(substring in str(row) for substring in ['marketing','pr','truyền thông']):
            danh_tieng_cong_ty_4.append(row)

        if any(substring in str(row) for substring in ['phát triển bền vững', 'phát triển', 'bền vững']):
            phat_trien_ben_vung.append(row)

        if any(substring in str(row) for substring in ['tôn trọng','thái độ','lắng nghe']):
            lanh_dao_quan_ly_1.append(row)
        if any(substring in str(row) for substring in ['định hướng','lộ trình thăng tiến','thăng tiến']):
            lanh_dao_quan_ly_2.append(row)
        if any(substring in str(row) for substring in ['tầm nhìn','chiến lược']):
            lanh_dao_quan_ly_3.append(row)
        if any(substring in str(row) for substring in ['phân biệt','động viên','công bằng','minh bạch']):
            lanh_dao_quan_ly_4.append(row)

        if any(substring in str(row) for substring in ['đa dạng', 'công bằng', 'hòa nhập']):
            da_dang_cong_bang_hoa_nhap.append(row)

        if any(substring in str(row) for substring in ['lương']):
            luong_phuc_loi_1.append(row)
        if any(substring in str(row) for substring in ['thưởng']):
            luong_phuc_loi_2.append(row)
        if any(substring in str(row) for substring in ['phúc lợi','chính sách']):
            luong_phuc_loi_3.append(row)

        if any(substring in str(row) for substring in ['sức khỏe']):
            cong_viec_suc_khoe_toan_dien_1.append(row)
        if any(substring in str(row) for substring in ['thời gian']):
            cong_viec_suc_khoe_toan_dien_2.append(row)

        if any(substring in str(row) for substring in ['phát triển nghề nghiệp']):
            co_hoi_dao_tao_phat_trien_1.append(row)
        if any(substring in str(row) for substring in ['chương trình tcc']):
            co_hoi_dao_tao_phat_trien_2.append(row)

        if any(substring in str(row) for substring in ['teambuilding','yep']):
            muc_do_gan_bo_1.append(row)
        if any(substring in str(row) for substring in ['outrace','hội thao']):
            muc_do_gan_bo_2.append(row)
        if any(substring in str(row) for substring in ['hoạt động gắn kết']):
            muc_do_gan_bo_3.append(row)

        if any(substring in str(row) for substring in ['thủ tục','giấy tờ']):
            cong_viec_co_cau_quy_trinh_1.append(row)
        if any(substring in str(row) for substring in ['cấu trúc','cơ cấu']):
            cong_viec_co_cau_quy_trinh_2.append(row)


# execute label
for cmt in moi_truong_lam_viec_1:
    cmt_rep = ana_sent(cmt)
    moi_truong_lam_viec_1_reply.append(cmt_rep)
for cmt in moi_truong_lam_viec_2:
    cmt_rep = ana_sent(cmt)
    moi_truong_lam_viec_2_reply.append(cmt_rep)
for cmt in moi_truong_lam_viec_3:
    cmt_rep = ana_sent(cmt)
    moi_truong_lam_viec_3_reply.append(cmt_rep)

for cmt in danh_tieng_cong_ty_1:
    cmt_rep = ana_sent(cmt)
    danh_tieng_cong_ty_1_reply.append(cmt_rep)
for cmt in danh_tieng_cong_ty_2:
    cmt_rep = ana_sent(cmt)
    danh_tieng_cong_ty_2_reply.append(cmt_rep)
for cmt in danh_tieng_cong_ty_3:
    cmt_rep = ana_sent(cmt)
    danh_tieng_cong_ty_3_reply.append(cmt_rep)
for cmt in danh_tieng_cong_ty_4:
    cmt_rep = ana_sent(cmt)
    danh_tieng_cong_ty_4_reply.append(cmt_rep)

for cmt in phat_trien_ben_vung:
    cmt_rep = ana_sent(cmt)
    phat_trien_ben_vung_reply.append(cmt_rep)

for cmt in lanh_dao_quan_ly_1:
    cmt_rep = ana_sent(cmt)
    lanh_dao_quan_ly_1_reply.append(cmt_rep)
for cmt in lanh_dao_quan_ly_2:
    cmt_rep = ana_sent(cmt)
    lanh_dao_quan_ly_2_reply.append(cmt_rep)
for cmt in lanh_dao_quan_ly_3:
    cmt_rep = ana_sent(cmt)
    lanh_dao_quan_ly_3_reply.append(cmt_rep)
for cmt in lanh_dao_quan_ly_4:
    cmt_rep = ana_sent(cmt)
    lanh_dao_quan_ly_4_reply.append(cmt_rep)

for cmt in da_dang_cong_bang_hoa_nhap:
    cmt_rep = ana_sent(cmt)
    da_dang_cong_bang_hoa_nhap_reply.append(cmt_rep)

for cmt in luong_phuc_loi_1:
    cmt_rep = ana_sent(cmt)
    luong_phuc_loi_1_reply.append(cmt_rep)
for cmt in luong_phuc_loi_2:
    cmt_rep = ana_sent(cmt)
    luong_phuc_loi_2_reply.append(cmt_rep)
for cmt in luong_phuc_loi_3:
    cmt_rep = ana_sent(cmt)
    luong_phuc_loi_3_reply.append(cmt_rep)

for cmt in cong_viec_suc_khoe_toan_dien_1:
    cmt_rep = ana_sent(cmt)
    cong_viec_suc_khoe_toan_dien_1_reply.append(cmt_rep)
for cmt in cong_viec_suc_khoe_toan_dien_2:
    cmt_rep = ana_sent(cmt)
    cong_viec_suc_khoe_toan_dien_2_reply.append(cmt_rep)

for cmt in co_hoi_dao_tao_phat_trien_1:
    cmt_rep = ana_sent(cmt)
    co_hoi_dao_tao_phat_trien_1_reply.append(cmt_rep)
for cmt in co_hoi_dao_tao_phat_trien_2:
    cmt_rep = ana_sent(cmt)
    co_hoi_dao_tao_phat_trien_2_reply.append(cmt_rep)

for cmt in muc_do_gan_bo_1:
    cmt_rep = ana_sent(cmt)
    muc_do_gan_bo_1_reply.append(cmt_rep)
for cmt in muc_do_gan_bo_2:
    cmt_rep = ana_sent(cmt)
    muc_do_gan_bo_2_reply.append(cmt_rep)
for cmt in muc_do_gan_bo_3:
    cmt_rep = ana_sent(cmt)
    muc_do_gan_bo_3_reply.append(cmt_rep)

for cmt in cong_viec_co_cau_quy_trinh_1:
    cmt_rep = ana_sent(cmt)
    cong_viec_co_cau_quy_trinh_1_reply.append(cmt_rep)
for cmt in cong_viec_co_cau_quy_trinh_1:
    cmt_rep = ana_sent(cmt)
    cong_viec_co_cau_quy_trinh_2_reply.append(cmt_rep)

# systemize dataframe
df_moi_truong_lam_viec_1 = pd.DataFrame((moi_truong_lam_viec_1, moi_truong_lam_viec_1_reply)).transpose().rename(columns={0 : 'mạng, internet, wifi, cơ sở hạ tầng cntt, hệ thống', 1: 'reply'})
df_moi_truong_lam_viec_2 = pd.DataFrame((moi_truong_lam_viec_2, moi_truong_lam_viec_2_reply)).transpose().rename(columns={0 : 'trang thiết bị, dụng cụ, laptop', 1: 'reply'})
df_moi_truong_lam_viec_3 = pd.DataFrame((moi_truong_lam_viec_3, moi_truong_lam_viec_3_reply)).transpose().rename(columns={0 : 'bàn làm việc, không gian tiện ích, phòng họp, cơ sở vật chất, văn phòng phẩm', 1: 'reply'})

df_danh_tieng_cong_ty_1 = pd.DataFrame((danh_tieng_cong_ty_1, danh_tieng_cong_ty_1_reply)).transpose().rename(columns={0 : 'sản phẩm, thương hiệu, sản phẩm food, thực phẩm, con giống', 1: 'reply'})
df_danh_tieng_cong_ty_2 = pd.DataFrame((danh_tieng_cong_ty_2, danh_tieng_cong_ty_2_reply)).transpose().rename(columns={0 : 'chất lượng', 1: 'reply'})
df_danh_tieng_cong_ty_3 = pd.DataFrame((danh_tieng_cong_ty_3, danh_tieng_cong_ty_3_reply)).transpose().rename(columns={0 : 'khách hàng, đối tác', 1: 'reply'})
df_danh_tieng_cong_ty_4 = pd.DataFrame((danh_tieng_cong_ty_4, danh_tieng_cong_ty_4_reply)).transpose().rename(columns={0 : 'marketing, pr, truyền thông', 1: 'reply'})

df_phat_trien_ben_vung = pd.DataFrame((phat_trien_ben_vung, phat_trien_ben_vung_reply)).transpose().rename(columns={0 : 'phát triển bền vững, phát triển, bền vững', 1: 'reply'})

df_lanh_dao_quan_ly_1 = pd.DataFrame((lanh_dao_quan_ly_1, lanh_dao_quan_ly_1_reply)).transpose().rename(columns={0 : 'tôn trọng, thái độ, lắng nghe', 1: 'reply'})
df_lanh_dao_quan_ly_2 = pd.DataFrame((lanh_dao_quan_ly_2, lanh_dao_quan_ly_2_reply)).transpose().rename(columns={0 : 'định hướng, lộ trình thăng tiến, thăng tiến', 1: 'reply'})
df_lanh_dao_quan_ly_3 = pd.DataFrame((lanh_dao_quan_ly_3, lanh_dao_quan_ly_3_reply)).transpose().rename(columns={0 : 'tầm nhìn, chiến lược', 1: 'reply'})
df_lanh_dao_quan_ly_4 = pd.DataFrame((lanh_dao_quan_ly_4, lanh_dao_quan_ly_4_reply)).transpose().rename(columns={0 : 'phân biệt, động viên, công bằng, minh bạch', 1: 'reply'})

df_da_dang_cong_bang_hoa_nhap = pd.DataFrame((da_dang_cong_bang_hoa_nhap, da_dang_cong_bang_hoa_nhap_reply)).transpose().rename(columns={0 : 'đa dạng, công bằng, hòa nhập', 1: 'reply'})

df_luong_phuc_loi_1 = pd.DataFrame((luong_phuc_loi_1, luong_phuc_loi_1_reply)).transpose().rename(columns={0 : 'lương', 1: 'reply'})
df_luong_phuc_loi_2 = pd.DataFrame((luong_phuc_loi_2, luong_phuc_loi_2_reply)).transpose().rename(columns={0 : 'thưởng', 1: 'reply'})
df_luong_phuc_loi_3 = pd.DataFrame((luong_phuc_loi_3, luong_phuc_loi_3_reply)).transpose().rename(columns={0 : 'phúc lợi, chính sách', 1: 'reply'})

df_cong_viec_suc_khoe_toan_dien_1 = pd.DataFrame((cong_viec_suc_khoe_toan_dien_1, cong_viec_suc_khoe_toan_dien_1_reply)).transpose().rename(columns={0 : 'sức khỏe', 1: 'reply'})
df_cong_viec_suc_khoe_toan_dien_2 = pd.DataFrame((cong_viec_suc_khoe_toan_dien_2, cong_viec_suc_khoe_toan_dien_2_reply)).transpose().rename(columns={0 : 'thời gian', 1: 'reply'})

df_co_hoi_dao_tao_phat_trien_1 = pd.DataFrame((co_hoi_dao_tao_phat_trien_1, co_hoi_dao_tao_phat_trien_1_reply)).transpose().rename(columns={0 : 'phát triển nghề nghiệp', 1: 'reply'})
df_co_hoi_dao_tao_phat_trien_2 = pd.DataFrame((co_hoi_dao_tao_phat_trien_2, co_hoi_dao_tao_phat_trien_2_reply)).transpose().rename(columns={0 : 'chương trình tcc', 1: 'reply'})

df_muc_do_gan_bo_1 = pd.DataFrame((muc_do_gan_bo_1, muc_do_gan_bo_1_reply)).transpose().rename(columns={0 : 'teambuilding, yep', 1: 'reply'})
df_muc_do_gan_bo_2 = pd.DataFrame((muc_do_gan_bo_2, muc_do_gan_bo_2_reply)).transpose().rename(columns={0 : 'outrace, hội thao', 1: 'reply'})
df_muc_do_gan_bo_3 = pd.DataFrame((muc_do_gan_bo_3, muc_do_gan_bo_3_reply)).transpose().rename(columns={0 : 'hoạt động gắn kết', 1: 'reply'})

df_cong_viec_co_cau_quy_trinh_1 = pd.DataFrame((cong_viec_co_cau_quy_trinh_1, cong_viec_co_cau_quy_trinh_1_reply)).transpose().rename(columns={0 : 'thủ tục, giấy tờ', 1: 'reply'})
df_cong_viec_co_cau_quy_trinh_2 = pd.DataFrame((cong_viec_co_cau_quy_trinh_2, cong_viec_co_cau_quy_trinh_2_reply)).transpose().rename(columns={0 : 'cấu trúc, cơ cấu', 1: 'reply'})


# execute export as excel
#with pd.ExcelWriter('comment_bu_vn.xlsx') as writer:
with pd.ExcelWriter('comment_bu_over_vn.xlsx') as writer:

    df_moi_truong_lam_viec_1.to_excel(writer, sheet_name = 'moi_truong_lam_viec')
    df_moi_truong_lam_viec_2.to_excel(writer, sheet_name = 'moi_truong_lam_viec', startrow = len(df_moi_truong_lam_viec_1) + 3)
    df_moi_truong_lam_viec_3.to_excel(writer, sheet_name = 'moi_truong_lam_viec', startrow = len(df_moi_truong_lam_viec_1) + 3 + len(df_moi_truong_lam_viec_2) + 3)

    df_danh_tieng_cong_ty_1.to_excel(writer, sheet_name = 'danh_tieng_cong_ty')
    df_danh_tieng_cong_ty_2.to_excel(writer, sheet_name = 'danh_tieng_cong_ty', startrow = len(df_danh_tieng_cong_ty_1) + 3)
    df_danh_tieng_cong_ty_3.to_excel(writer, sheet_name = 'danh_tieng_cong_ty', startrow = len(df_danh_tieng_cong_ty_1) + 3 + len(df_danh_tieng_cong_ty_2) + 3)
    df_danh_tieng_cong_ty_4.to_excel(writer, sheet_name = 'danh_tieng_cong_ty', startrow = len(df_danh_tieng_cong_ty_1) + 3 + len(df_danh_tieng_cong_ty_2) + 3 + len(df_danh_tieng_cong_ty_3) + 3)

    df_phat_trien_ben_vung.to_excel(writer, sheet_name = 'phat_trien_ben_vung')

    df_lanh_dao_quan_ly_1.to_excel(writer, sheet_name = 'lanh_dao_quan_ly')
    df_lanh_dao_quan_ly_2.to_excel(writer, sheet_name = 'lanh_dao_quan_ly', startrow = len(df_lanh_dao_quan_ly_1) + 3)
    df_lanh_dao_quan_ly_3.to_excel(writer, sheet_name = 'lanh_dao_quan_ly', startrow = len(df_lanh_dao_quan_ly_1) + 3 + len(df_lanh_dao_quan_ly_2) + 3)
    df_lanh_dao_quan_ly_4.to_excel(writer, sheet_name = 'lanh_dao_quan_ly', startrow = len(df_lanh_dao_quan_ly_1) + 3 + len(df_lanh_dao_quan_ly_2) + 3 + len(df_lanh_dao_quan_ly_3) + 3)

    df_da_dang_cong_bang_hoa_nhap.to_excel(writer, sheet_name = 'da_dang_cong_bang_hoa_nhap')

    df_luong_phuc_loi_1.to_excel(writer, sheet_name = 'luong_phuc_loi')
    df_luong_phuc_loi_2.to_excel(writer, sheet_name = 'luong_phuc_loi', startrow = len(df_luong_phuc_loi_1) + 3)
    df_luong_phuc_loi_3.to_excel(writer, sheet_name = 'luong_phuc_loi', startrow = len(df_luong_phuc_loi_1) + 3 + len(df_luong_phuc_loi_2) + 3)

    df_cong_viec_suc_khoe_toan_dien_1.to_excel(writer, sheet_name = 'cong_viec_suc_khoe_toan_dien')
    df_cong_viec_suc_khoe_toan_dien_2.to_excel(writer, sheet_name = 'cong_viec_suc_khoe_toan_dien', startrow = len(df_cong_viec_suc_khoe_toan_dien_1) + 3)

    df_co_hoi_dao_tao_phat_trien_1.to_excel(writer, sheet_name = 'co_hoi_dao_tao_phat_trien')
    df_co_hoi_dao_tao_phat_trien_2.to_excel(writer, sheet_name = 'co_hoi_dao_tao_phat_trien', startrow = len(df_co_hoi_dao_tao_phat_trien_1) + 3)

    df_muc_do_gan_bo_1.to_excel(writer, sheet_name = 'muc_do_gan_bo')
    df_muc_do_gan_bo_2.to_excel(writer, sheet_name = 'muc_do_gan_bo', startrow = len(df_muc_do_gan_bo_1) + 3)
    df_muc_do_gan_bo_3.to_excel(writer, sheet_name = 'muc_do_gan_bo', startrow = len(df_muc_do_gan_bo_1) + 3 + len(df_muc_do_gan_bo_2) + 3)

    df_cong_viec_co_cau_quy_trinh_1.to_excel(writer, sheet_name = 'cong_viec_co_cau_quy_trinh')
    df_cong_viec_co_cau_quy_trinh_2.to_excel(writer, sheet_name = 'cong_viec_co_cau_quy_trinh', startrow = len(df_cong_viec_co_cau_quy_trinh_1) + 3)

