import pandas as pd

def load_input_data(filename="積算入力シート.xlsx"):
    # シート1の読み込み
    df = pd.read_excel(filename, sheet_name='基本情報・フリーザー仕様')
    
    # 項目名をキーにした辞書に変換して扱いやすくする
    input_data = dict(zip(df['項目名'], df['入力欄']))
    return input_data

def calculate_freezer_costs(data):
    print(f"\n=== 📊 積算計算シミュレーション: {data['プロジェクト名']} ===")
    
    # ---------------------------------------------------------
    # 1. パネル面積・防熱工事の計算
    # ---------------------------------------------------------
    width = float(data['機巾 [mm]']) / 1000
    length = float(data['機長 [mm]']) / 1000
    height = float(data['機高 [mm]']) / 1000
    
    # 表面積 = (幅×長)×2(天底) + (幅×高)×2(前後) + (長×高)×2(側面)
    total_area = (width * length * 2) + (width * height * 2) + (length * height * 2)
    
    # 扉面積の控除
    door_qty = float(data['扉の枚数 [枚]'])
    door_width = float(data['扉の幅 [mm]']) / 1000
    door_height = float(data['扉の高さ [mm]']) / 1000
    door_area = door_width * door_height * door_qty
    
    net_panel_area = total_area - door_area
    
    print(f"☑️ フリーザー総表面積: {total_area:.2f} ㎡")
    print(f"☑️ 扉面積の控除: -{door_area:.2f} ㎡")
    print(f"✅ 実質パネル施工面積: {net_panel_area:.2f} ㎡")

    # ---------------------------------------------------------
    # 2. 冷却水配管・ポンプの計算（水冷式の場合）
    # ---------------------------------------------------------
    condenser_type = str(data['凝縮方式']).strip()
    print(f"\n☑️ 凝縮方式: {condenser_type}")
    
    if condenser_type == "水冷式":
        straight_length = float(data['配管直管距離(往復) [m]'])
        complexity = str(data['配管ルート複雑さ']).strip()
        
        # 複雑さによる割増係数（実績データに基づいて後から調整可能）
        multiplier_map = {"シンプル": 1.3, "普通": 1.5, "複雑": 2.0}
        multiplier = multiplier_map.get(complexity, 1.5) # デフォルトは普通(1.5)
        
        equivalent_length = straight_length * multiplier
        head_diff = float(data['高低差(実揚程) [m]'])
        
        print(f"✅ 冷却水配管 直管距離: {straight_length}m -> 全相当長(継手加味): {equivalent_length:.1f}m (係数:{multiplier})")
        print(f"✅ ポンプ計算用 高低差: {head_diff}m")
        print(f"   ※この相当長を用いて、マニュアルDBから最適口径の配管単価×距離を抽出します。")
    else:
        print("✅ 空冷式のため、冷却塔・冷却水ポンプの積算はスキップします。")

if __name__ == "__main__":
    try:
        input_data = load_input_data()
        calculate_freezer_costs(input_data)
    except FileNotFoundError:
        print("エラー: 先に create_template.py を実行して '積算入力シート.xlsx' を作成してください。")
