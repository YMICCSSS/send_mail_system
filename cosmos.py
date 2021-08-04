from azure.cosmos import  CosmosClient, PartitionKey
import uuid
import wmi



# 填寫你們自己的cosmos端點及金鑰
endpoint = "XXXX"
key = 'XXXXX'

# 查詢機器代碼的物件
c = wmi.WMI()

client = CosmosClient(endpoint, key)
my_uuid = uuid.uuid4()
result = {}


# 產生授權碼，需填入公司名稱，記錄是授權哪間公司使用,並設定該公司的授權到期日
def creat_authorization_code(company_name,authorization_day):
    # 資料庫
    database_name = 'Send_email'
    database = client.create_database_if_not_exists(id=database_name)

    # 容器
    container_name = 'Send_email_authorization'

    # 整合資料
    result["Send_email_authorization"] = str(uuid.uuid4())
    result["company_name"] = company_name

    # 授權時期
    result["authorization_day"] = authorization_day

    # 加入一個不重複id,儲存此次事件資料
    result["id"] = str(uuid.uuid4())

    # container = 紀錄客戶 follow 時的資料庫容器
    container = database.create_container_if_not_exists(
        id=container_name,
        partition_key=PartitionKey(path="/id"),
        offer_throughput=400
    )

    container.create_item(body=result)

    print("創建成功")

# 查看授權資料庫中是否有此授權碼，如果有就回傳這筆資料
def query_authorization_code(Send_email_authorization):
    # 資料庫
    database_name = 'Send_email'
    database = client.create_database_if_not_exists(id=database_name)

    # 容器
    container_name = 'Send_email_authorization'

    container = database.create_container_if_not_exists(
        id=container_name,
        partition_key=PartitionKey(path="/id"),
        offer_throughput=400
    )

    query = "SELECT * FROM c WHERE c.Send_email_authorization IN ( '%s' )" % (Send_email_authorization)
    items = list(container.query_items(
        query=query,
        enable_cross_partition_query=True
    ))

    return items

# 第一次使用授權碼時毀啟動這個函式，將授權碼存到「已使用授權碼」資料庫，代表授權碼已被啟用
def used_authorization_code(acad_authorization,company_name,authorization_end_day,authorization_end_day_ts):
    # 資料庫
    database_name = 'Send_email'
    database = client.create_database_if_not_exists(id=database_name)

    # 容器
    container_name = 'used_authorization_code'

    # 整合資料
    result["Send_email_authorization"] = acad_authorization

    # 將主機板序列號寫進資料庫
    result["cpu_serial_number"] = c.Win32_BaseBoard()[0].SerialNumber.strip()

    # 公司、使用者名稱
    result["company_name"] = company_name

    # 到期日
    result["authorization_end_day"] = authorization_end_day

    # 到期日的時間截記
    result["authorization_end_day_ts"] = authorization_end_day_ts


    # 加入一個不重複id,儲存此次事件資料
    result["id"] = str(uuid.uuid4())

    # container = 紀錄客戶 follow 時的資料庫容器
    container = database.create_container_if_not_exists(
        id=container_name,
        partition_key=PartitionKey(path="/id"),
        offer_throughput=400
    )

    container.create_item(body=result)


# 查看授權碼是否已已在「已使用授權碼」資料庫中，如果有就回傳這筆資料
def query_used_authorization_code(Send_email_authorization):
    # 資料庫
    database_name = 'Send_email'
    database = client.create_database_if_not_exists(id=database_name)

    # 容器
    container_name = 'used_authorization_code'

    container = database.create_container_if_not_exists(
        id=container_name,
        partition_key=PartitionKey(path="/id"),
        offer_throughput=400
    )

    query = "SELECT * FROM c WHERE c.Send_email_authorization IN ( '%s' )" % (Send_email_authorization)
    items = list(container.query_items(
        query=query,
        enable_cross_partition_query=True
    ))

    return items

# 紀錄使用次數
def save_frequency(frequency,company_name):
    # 資料庫
    database_name = 'Send_email'
    database = client.create_database_if_not_exists(id=database_name)

    # 容器
    container_name = 'frequency'

    #整合資料
    result["frequency"] = frequency

    #公司名稱
    result["company_name"] = company_name

    # 加入一個不重複id,儲存此次事件資料
    result["id"] = str(uuid.uuid4())

    # container = 紀錄客戶 follow 時的資料庫容器
    container = database.create_container_if_not_exists(
        id=container_name,
        partition_key=PartitionKey(path="/id"),
        offer_throughput=400
    )

    container.create_item(body=result)
