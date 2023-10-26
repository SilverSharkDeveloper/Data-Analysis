### Data-Analysis Side Project
### RFM 분석
####  사용자별로 얼마나 최근에, 얼마나 자주, 얼마나 많은 금액을 지출했는지에 따라 사용자들의 분포를 확인할 수 있고, 사용자 그룹(또는 등급)을 나누어 분류하는 분석 기법이다. 구매 가능성이 높은 고객을 선정할 때 용이한 데이터 분석 방법이며, 사용자들의 평소 구매 패턴을 기준으로 분류를 진행하기 때문에 각 사용자 그룹의 특성에 따라 차별화된 마케팅 메세지를 전달할 수 있다.

- Recency : 얼마나 최근에 구매했는가
- Frequency : 얼마나 자주 구매했는가
- Monetary : 얼마나 많은 금액을 지출했는가

### 🛒이커머스 플랫폼 A기업의 RFM 분석
📌이커머스: 온라인을 통해 상품이나 서비스를 사고파는 서비스 (쿠팡, 11번가, 네이버 등)

 
<table style="width: 50%; margin-left:10px;">
    <caption>고객 분석</caption>
    <tr>
        <th>사용자</th>
        <th>구매 횟수</th>
        <th>구매 금액</th>
        <th>최근 구매일</th>
    </tr>
    <tr>
        <th>한동석</th>
        <th>45</th>
        <th>1,980,000</th>
        <th>2개월 전</th>
    </tr>
    <tr>
        <th>주선유</th>
        <th>2</th>
        <th>45,320</th>
        <th>1년 전</th>
    </tr>
</table>

##### 👓 "한동석" 고객을 VIP로 선정해서 연말 선물을 전달하면, 충성심있는 고객으로 유지할 수 있는 전략을 세울 수도 있고,<br> 🎫 "주선유" 고객에게 할인 쿠폰 등 자사의 플랫폼을 이용할 거리를 전달함으로써, 구매를 유도할 수 있는 전략을 세울 수도 있다.

####  RFM을 사용하면 사용자의 특성별로 각기 다른 정책을 적용하고 서비스를 더 잘 사용하게끔 유도하는 전략을 세워볼 수가 있다.

### 데이터셋 확인



```python
import pandas as pd
#엑셀 또는 csv파일 불러오기(csv 파일을 읽어올 때는 구분점이 무엇인지 판단해서 작성해주어여함 콤마(,)가 디폴트
#액셀은 이미 구분되어 있으므로 적을필요 없음 )
retail_orders_df = pd.read_excel('./datasets/Retailer Sales Orders - EU-UK.xlsx')
display(retail_orders_df)

```


<div>

<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Order</th>
      <th>Item Name</th>
      <th>Quantity</th>
      <th>Date</th>
      <th>Time</th>
      <th>Price</th>
      <th>Customer ID</th>
      <th>Country</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>540040</td>
      <td>FELT FARM ANIMAL WHITE BUNNY</td>
      <td>48</td>
      <td>2011-01-04 14:20:00</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.19</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>1</th>
      <td>540040</td>
      <td>EAU DE NIL LOVE BIRD CANDLE</td>
      <td>24</td>
      <td>2011-01-04 14:20:00</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.42</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>2</th>
      <td>540040</td>
      <td>IVORY LOVE BIRD CANDLE</td>
      <td>24</td>
      <td>2011-01-04 14:20:00</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.42</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>3</th>
      <td>540040</td>
      <td>BLACK LOVE BIRD CANDLE</td>
      <td>24</td>
      <td>2011-01-04 14:20:00</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.42</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>4</th>
      <td>540040</td>
      <td>FELTCRAFT DOLL ROSIE</td>
      <td>12</td>
      <td>2011-01-04 14:20:00</td>
      <td>2011-01-04 14:20:00</td>
      <td>2.95</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>...</th>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
    </tr>
    <tr>
      <th>470645</th>
      <td>581585</td>
      <td>FAIRY TALE COTTAGE NIGHT LIGHT</td>
      <td>12</td>
      <td>2011-12-09 12:31:00</td>
      <td>2011-12-09 12:31:00</td>
      <td>1.95</td>
      <td>15804.0</td>
      <td>United Kingdom</td>
    </tr>
    <tr>
      <th>470646</th>
      <td>581586</td>
      <td>LARGE CAKE STAND  HANGING STRAWBERY</td>
      <td>8</td>
      <td>2011-12-09 12:49:00</td>
      <td>2011-12-09 12:49:00</td>
      <td>2.95</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
    </tr>
    <tr>
      <th>470647</th>
      <td>581586</td>
      <td>SET OF 3 HANGING OWLS OLLIE BEAK</td>
      <td>24</td>
      <td>2011-12-09 12:49:00</td>
      <td>2011-12-09 12:49:00</td>
      <td>1.25</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
    </tr>
    <tr>
      <th>470648</th>
      <td>581586</td>
      <td>RED RETROSPOT ROUND CAKE TINS</td>
      <td>24</td>
      <td>2011-12-09 12:49:00</td>
      <td>2011-12-09 12:49:00</td>
      <td>8.95</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
    </tr>
    <tr>
      <th>470649</th>
      <td>581586</td>
      <td>DOORMAT RED RETROSPOT</td>
      <td>10</td>
      <td>2011-12-09 12:49:00</td>
      <td>2011-12-09 12:49:00</td>
      <td>7.08</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
    </tr>
  </tbody>
</table>
<p>470650 rows × 8 columns</p>
</div>


### 컬럼별 갯수, 평균, 최솟값, 4분위분포, 최댓값, 표준편차 확인하기



```python
#컬럼별 갯수, 평균, 최솟값, 4분위분포, 최댓값, 표준편차 확인하기
display(retail_orders_df.describe().T)
```


<div>

<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>count</th>
      <th>mean</th>
      <th>min</th>
      <th>25%</th>
      <th>50%</th>
      <th>75%</th>
      <th>max</th>
      <th>std</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>Order</th>
      <td>470650.0</td>
      <td>561860.118721</td>
      <td>539993.0</td>
      <td>550987.0</td>
      <td>562592.0</td>
      <td>572758.0</td>
      <td>581587.0</td>
      <td>12347.899552</td>
    </tr>
    <tr>
      <th>Quantity</th>
      <td>470650.0</td>
      <td>10.338173</td>
      <td>1.0</td>
      <td>1.0</td>
      <td>3.0</td>
      <td>10.0</td>
      <td>80995.0</td>
      <td>164.392408</td>
    </tr>
    <tr>
      <th>Date</th>
      <td>470650</td>
      <td>2011-07-22 11:00:08.260660736</td>
      <td>2011-01-04 10:00:00</td>
      <td>2011-04-21 18:10:00</td>
      <td>2011-08-07 15:43:00</td>
      <td>2011-10-25 18:39:00</td>
      <td>2011-12-09 12:50:00</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>Time</th>
      <td>470650</td>
      <td>2011-07-22 11:00:08.260660736</td>
      <td>2011-01-04 10:00:00</td>
      <td>2011-04-21 18:10:00</td>
      <td>2011-08-07 15:43:00</td>
      <td>2011-10-25 18:39:00</td>
      <td>2011-12-09 12:50:00</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>Price</th>
      <td>470650.0</td>
      <td>3.222389</td>
      <td>0.04</td>
      <td>1.25</td>
      <td>2.08</td>
      <td>4.13</td>
      <td>649.5</td>
      <td>4.335077</td>
    </tr>
    <tr>
      <th>Customer ID</th>
      <td>355713.0</td>
      <td>15346.934967</td>
      <td>12346.0</td>
      <td>14005.0</td>
      <td>15296.0</td>
      <td>16837.0</td>
      <td>18287.0</td>
      <td>1695.614811</td>
    </tr>
  </tbody>
</table>
</div>


### 컬럼의 자세한 정보 (null 개수, dtype)확인 및 메모리 사용량 확인하기


```python
print(retail_orders_df.info())
```

    <class 'pandas.core.frame.DataFrame'>
    RangeIndex: 470650 entries, 0 to 470649
    Data columns (total 8 columns):
     #   Column       Non-Null Count   Dtype         
    ---  ------       --------------   -----         
     0   Order        470650 non-null  int64         
     1   Item Name    470650 non-null  object        
     2   Quantity     470650 non-null  int64         
     3   Date         470650 non-null  datetime64[ns]
     4   Time         470650 non-null  datetime64[ns]
     5   Price        470650 non-null  float64       
     6   Customer ID  355713 non-null  float64       
     7   Country      470650 non-null  object        
    dtypes: datetime64[ns](2), float64(2), int64(2), object(2)
    memory usage: 28.7+ MB
    None
    

### 컬럼분석
> Order : 주문번호  
Item Name : 상품명  
Quantity : 주문 갯수  
Date : 주문 날짜  
Time : 주문 시간  
Price : 상품 가격  
Customer ID : 고객 번호  
Country : 주문 나라 

### 중복 데이터 처리


```python
#중복행 갯수 확인
retail_orders_df.duplicated().sum()

#중복행 제거 -> 파라미터에 컬럼명 전달 시 컬럼 값이 같으면 같다고 판단 아무것도 전달하지 않으면 모든 컬럼값이 같아야 같다고 판단
retail_orders_df.drop_duplicates(inplace=True)
display(retail_orders_df)

#인덱스 재생성 -> drop옵션에 true를 주면 기존 인덱스가 제거되고 false를 주면 기존인덱스가 컬럼으로 추가됨
retail_orders_df.reset_index(drop=True,inplace=True)
display(retail_orders_df)
```


<div>

<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Order</th>
      <th>Item Name</th>
      <th>Quantity</th>
      <th>Date</th>
      <th>Time</th>
      <th>Price</th>
      <th>Customer ID</th>
      <th>Country</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>540040</td>
      <td>FELT FARM ANIMAL WHITE BUNNY</td>
      <td>48</td>
      <td>2011-01-04 14:20:00</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.19</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>1</th>
      <td>540040</td>
      <td>EAU DE NIL LOVE BIRD CANDLE</td>
      <td>24</td>
      <td>2011-01-04 14:20:00</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.42</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>2</th>
      <td>540040</td>
      <td>IVORY LOVE BIRD CANDLE</td>
      <td>24</td>
      <td>2011-01-04 14:20:00</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.42</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>3</th>
      <td>540040</td>
      <td>BLACK LOVE BIRD CANDLE</td>
      <td>24</td>
      <td>2011-01-04 14:20:00</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.42</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>4</th>
      <td>540040</td>
      <td>FELTCRAFT DOLL ROSIE</td>
      <td>12</td>
      <td>2011-01-04 14:20:00</td>
      <td>2011-01-04 14:20:00</td>
      <td>2.95</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>...</th>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
    </tr>
    <tr>
      <th>470645</th>
      <td>581585</td>
      <td>FAIRY TALE COTTAGE NIGHT LIGHT</td>
      <td>12</td>
      <td>2011-12-09 12:31:00</td>
      <td>2011-12-09 12:31:00</td>
      <td>1.95</td>
      <td>15804.0</td>
      <td>United Kingdom</td>
    </tr>
    <tr>
      <th>470646</th>
      <td>581586</td>
      <td>LARGE CAKE STAND  HANGING STRAWBERY</td>
      <td>8</td>
      <td>2011-12-09 12:49:00</td>
      <td>2011-12-09 12:49:00</td>
      <td>2.95</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
    </tr>
    <tr>
      <th>470647</th>
      <td>581586</td>
      <td>SET OF 3 HANGING OWLS OLLIE BEAK</td>
      <td>24</td>
      <td>2011-12-09 12:49:00</td>
      <td>2011-12-09 12:49:00</td>
      <td>1.25</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
    </tr>
    <tr>
      <th>470648</th>
      <td>581586</td>
      <td>RED RETROSPOT ROUND CAKE TINS</td>
      <td>24</td>
      <td>2011-12-09 12:49:00</td>
      <td>2011-12-09 12:49:00</td>
      <td>8.95</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
    </tr>
    <tr>
      <th>470649</th>
      <td>581586</td>
      <td>DOORMAT RED RETROSPOT</td>
      <td>10</td>
      <td>2011-12-09 12:49:00</td>
      <td>2011-12-09 12:49:00</td>
      <td>7.08</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
    </tr>
  </tbody>
</table>
<p>465897 rows × 8 columns</p>
</div>



<div>

<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Order</th>
      <th>Item Name</th>
      <th>Quantity</th>
      <th>Date</th>
      <th>Time</th>
      <th>Price</th>
      <th>Customer ID</th>
      <th>Country</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>540040</td>
      <td>FELT FARM ANIMAL WHITE BUNNY</td>
      <td>48</td>
      <td>2011-01-04 14:20:00</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.19</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>1</th>
      <td>540040</td>
      <td>EAU DE NIL LOVE BIRD CANDLE</td>
      <td>24</td>
      <td>2011-01-04 14:20:00</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.42</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>2</th>
      <td>540040</td>
      <td>IVORY LOVE BIRD CANDLE</td>
      <td>24</td>
      <td>2011-01-04 14:20:00</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.42</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>3</th>
      <td>540040</td>
      <td>BLACK LOVE BIRD CANDLE</td>
      <td>24</td>
      <td>2011-01-04 14:20:00</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.42</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>4</th>
      <td>540040</td>
      <td>FELTCRAFT DOLL ROSIE</td>
      <td>12</td>
      <td>2011-01-04 14:20:00</td>
      <td>2011-01-04 14:20:00</td>
      <td>2.95</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>...</th>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
    </tr>
    <tr>
      <th>465892</th>
      <td>581585</td>
      <td>FAIRY TALE COTTAGE NIGHT LIGHT</td>
      <td>12</td>
      <td>2011-12-09 12:31:00</td>
      <td>2011-12-09 12:31:00</td>
      <td>1.95</td>
      <td>15804.0</td>
      <td>United Kingdom</td>
    </tr>
    <tr>
      <th>465893</th>
      <td>581586</td>
      <td>LARGE CAKE STAND  HANGING STRAWBERY</td>
      <td>8</td>
      <td>2011-12-09 12:49:00</td>
      <td>2011-12-09 12:49:00</td>
      <td>2.95</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
    </tr>
    <tr>
      <th>465894</th>
      <td>581586</td>
      <td>SET OF 3 HANGING OWLS OLLIE BEAK</td>
      <td>24</td>
      <td>2011-12-09 12:49:00</td>
      <td>2011-12-09 12:49:00</td>
      <td>1.25</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
    </tr>
    <tr>
      <th>465895</th>
      <td>581586</td>
      <td>RED RETROSPOT ROUND CAKE TINS</td>
      <td>24</td>
      <td>2011-12-09 12:49:00</td>
      <td>2011-12-09 12:49:00</td>
      <td>8.95</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
    </tr>
    <tr>
      <th>465896</th>
      <td>581586</td>
      <td>DOORMAT RED RETROSPOT</td>
      <td>10</td>
      <td>2011-12-09 12:49:00</td>
      <td>2011-12-09 12:49:00</td>
      <td>7.08</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
    </tr>
  </tbody>
</table>
<p>465897 rows × 8 columns</p>
</div>


### 결측치 처리


```python
#결측치 검사
retail_orders_df.isnull().sum()

#결측치 행 삭제
#결측치가 고객의 구분점인 CustomerID이므로 분석 후 연계 불가능.
# CustomerID가 있는 행만 분석하도록 한다

# retail_orders_df = retail_orders_df[~retail_orders_df.isnull()["Customer ID"]]
retail_orders_df.dropna(subset="Customer ID",inplace=True)
retail_orders_df.reset_index(inplace=True , drop=True)
retail_orders_df
```




<div>

<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Order</th>
      <th>Item Name</th>
      <th>Quantity</th>
      <th>Date</th>
      <th>Time</th>
      <th>Price</th>
      <th>Customer ID</th>
      <th>Country</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>540040</td>
      <td>FELT FARM ANIMAL WHITE BUNNY</td>
      <td>48</td>
      <td>2011-01-04 14:20:00</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.19</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>1</th>
      <td>540040</td>
      <td>EAU DE NIL LOVE BIRD CANDLE</td>
      <td>24</td>
      <td>2011-01-04 14:20:00</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.42</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>2</th>
      <td>540040</td>
      <td>IVORY LOVE BIRD CANDLE</td>
      <td>24</td>
      <td>2011-01-04 14:20:00</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.42</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>3</th>
      <td>540040</td>
      <td>BLACK LOVE BIRD CANDLE</td>
      <td>24</td>
      <td>2011-01-04 14:20:00</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.42</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>4</th>
      <td>540040</td>
      <td>FELTCRAFT DOLL ROSIE</td>
      <td>12</td>
      <td>2011-01-04 14:20:00</td>
      <td>2011-01-04 14:20:00</td>
      <td>2.95</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>...</th>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
    </tr>
    <tr>
      <th>351001</th>
      <td>581585</td>
      <td>FAIRY TALE COTTAGE NIGHT LIGHT</td>
      <td>12</td>
      <td>2011-12-09 12:31:00</td>
      <td>2011-12-09 12:31:00</td>
      <td>1.95</td>
      <td>15804.0</td>
      <td>United Kingdom</td>
    </tr>
    <tr>
      <th>351002</th>
      <td>581586</td>
      <td>LARGE CAKE STAND  HANGING STRAWBERY</td>
      <td>8</td>
      <td>2011-12-09 12:49:00</td>
      <td>2011-12-09 12:49:00</td>
      <td>2.95</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
    </tr>
    <tr>
      <th>351003</th>
      <td>581586</td>
      <td>SET OF 3 HANGING OWLS OLLIE BEAK</td>
      <td>24</td>
      <td>2011-12-09 12:49:00</td>
      <td>2011-12-09 12:49:00</td>
      <td>1.25</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
    </tr>
    <tr>
      <th>351004</th>
      <td>581586</td>
      <td>RED RETROSPOT ROUND CAKE TINS</td>
      <td>24</td>
      <td>2011-12-09 12:49:00</td>
      <td>2011-12-09 12:49:00</td>
      <td>8.95</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
    </tr>
    <tr>
      <th>351005</th>
      <td>581586</td>
      <td>DOORMAT RED RETROSPOT</td>
      <td>10</td>
      <td>2011-12-09 12:49:00</td>
      <td>2011-12-09 12:49:00</td>
      <td>7.08</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
    </tr>
  </tbody>
</table>
<p>351006 rows × 8 columns</p>
</div>



### 데이터 전처리

> Date 컬럼과 Time 컬럼이 중복된 컬럼이므로 Time컬럼 삭제  
Recency는 Date와 현재날짜를 기준으로 설정  
Frequency는 고객이 같은 주문빈도로 설정  
Monetary는 고객별 Price와 Quantity의 곱으로 설정  


```python
retail_orders_df.drop("Time",axis=1,inplace=True)
retail_orders_df
```




<div>

<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Order</th>
      <th>Item Name</th>
      <th>Quantity</th>
      <th>Date</th>
      <th>Price</th>
      <th>Customer ID</th>
      <th>Country</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>540040</td>
      <td>FELT FARM ANIMAL WHITE BUNNY</td>
      <td>48</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.19</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>1</th>
      <td>540040</td>
      <td>EAU DE NIL LOVE BIRD CANDLE</td>
      <td>24</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.42</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>2</th>
      <td>540040</td>
      <td>IVORY LOVE BIRD CANDLE</td>
      <td>24</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.42</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>3</th>
      <td>540040</td>
      <td>BLACK LOVE BIRD CANDLE</td>
      <td>24</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.42</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>4</th>
      <td>540040</td>
      <td>FELTCRAFT DOLL ROSIE</td>
      <td>12</td>
      <td>2011-01-04 14:20:00</td>
      <td>2.95</td>
      <td>12483.0</td>
      <td>Sweden</td>
    </tr>
    <tr>
      <th>...</th>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
    </tr>
    <tr>
      <th>351001</th>
      <td>581585</td>
      <td>FAIRY TALE COTTAGE NIGHT LIGHT</td>
      <td>12</td>
      <td>2011-12-09 12:31:00</td>
      <td>1.95</td>
      <td>15804.0</td>
      <td>United Kingdom</td>
    </tr>
    <tr>
      <th>351002</th>
      <td>581586</td>
      <td>LARGE CAKE STAND  HANGING STRAWBERY</td>
      <td>8</td>
      <td>2011-12-09 12:49:00</td>
      <td>2.95</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
    </tr>
    <tr>
      <th>351003</th>
      <td>581586</td>
      <td>SET OF 3 HANGING OWLS OLLIE BEAK</td>
      <td>24</td>
      <td>2011-12-09 12:49:00</td>
      <td>1.25</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
    </tr>
    <tr>
      <th>351004</th>
      <td>581586</td>
      <td>RED RETROSPOT ROUND CAKE TINS</td>
      <td>24</td>
      <td>2011-12-09 12:49:00</td>
      <td>8.95</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
    </tr>
    <tr>
      <th>351005</th>
      <td>581586</td>
      <td>DOORMAT RED RETROSPOT</td>
      <td>10</td>
      <td>2011-12-09 12:49:00</td>
      <td>7.08</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
    </tr>
  </tbody>
</table>
<p>351006 rows × 7 columns</p>
</div>



### Recency구하기


```python
from datetime import datetime
now = datetime.now()
retail_orders_df["Recency"] = (now - retail_orders_df["Date"]).dt.days
retail_orders_df

```




<div>

<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Order</th>
      <th>Item Name</th>
      <th>Quantity</th>
      <th>Date</th>
      <th>Price</th>
      <th>Customer ID</th>
      <th>Country</th>
      <th>Recency</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>540040</td>
      <td>FELT FARM ANIMAL WHITE BUNNY</td>
      <td>48</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.19</td>
      <td>12483.0</td>
      <td>Sweden</td>
      <td>4657</td>
    </tr>
    <tr>
      <th>1</th>
      <td>540040</td>
      <td>EAU DE NIL LOVE BIRD CANDLE</td>
      <td>24</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.42</td>
      <td>12483.0</td>
      <td>Sweden</td>
      <td>4657</td>
    </tr>
    <tr>
      <th>2</th>
      <td>540040</td>
      <td>IVORY LOVE BIRD CANDLE</td>
      <td>24</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.42</td>
      <td>12483.0</td>
      <td>Sweden</td>
      <td>4657</td>
    </tr>
    <tr>
      <th>3</th>
      <td>540040</td>
      <td>BLACK LOVE BIRD CANDLE</td>
      <td>24</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.42</td>
      <td>12483.0</td>
      <td>Sweden</td>
      <td>4657</td>
    </tr>
    <tr>
      <th>4</th>
      <td>540040</td>
      <td>FELTCRAFT DOLL ROSIE</td>
      <td>12</td>
      <td>2011-01-04 14:20:00</td>
      <td>2.95</td>
      <td>12483.0</td>
      <td>Sweden</td>
      <td>4657</td>
    </tr>
    <tr>
      <th>...</th>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
    </tr>
    <tr>
      <th>351001</th>
      <td>581585</td>
      <td>FAIRY TALE COTTAGE NIGHT LIGHT</td>
      <td>12</td>
      <td>2011-12-09 12:31:00</td>
      <td>1.95</td>
      <td>15804.0</td>
      <td>United Kingdom</td>
      <td>4318</td>
    </tr>
    <tr>
      <th>351002</th>
      <td>581586</td>
      <td>LARGE CAKE STAND  HANGING STRAWBERY</td>
      <td>8</td>
      <td>2011-12-09 12:49:00</td>
      <td>2.95</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
      <td>4318</td>
    </tr>
    <tr>
      <th>351003</th>
      <td>581586</td>
      <td>SET OF 3 HANGING OWLS OLLIE BEAK</td>
      <td>24</td>
      <td>2011-12-09 12:49:00</td>
      <td>1.25</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
      <td>4318</td>
    </tr>
    <tr>
      <th>351004</th>
      <td>581586</td>
      <td>RED RETROSPOT ROUND CAKE TINS</td>
      <td>24</td>
      <td>2011-12-09 12:49:00</td>
      <td>8.95</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
      <td>4318</td>
    </tr>
    <tr>
      <th>351005</th>
      <td>581586</td>
      <td>DOORMAT RED RETROSPOT</td>
      <td>10</td>
      <td>2011-12-09 12:49:00</td>
      <td>7.08</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
      <td>4318</td>
    </tr>
  </tbody>
</table>
<p>351006 rows × 8 columns</p>
</div>



### Monetary구하기


```python
retail_orders_df["Monetary"] =retail_orders_df["Quantity"] * retail_orders_df["Price"]
retail_orders_df
```




<div>

<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Order</th>
      <th>Item Name</th>
      <th>Quantity</th>
      <th>Date</th>
      <th>Price</th>
      <th>Customer ID</th>
      <th>Country</th>
      <th>Recency</th>
      <th>Monetary</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>540040</td>
      <td>FELT FARM ANIMAL WHITE BUNNY</td>
      <td>48</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.19</td>
      <td>12483.0</td>
      <td>Sweden</td>
      <td>4657</td>
      <td>9.12</td>
    </tr>
    <tr>
      <th>1</th>
      <td>540040</td>
      <td>EAU DE NIL LOVE BIRD CANDLE</td>
      <td>24</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.42</td>
      <td>12483.0</td>
      <td>Sweden</td>
      <td>4657</td>
      <td>10.08</td>
    </tr>
    <tr>
      <th>2</th>
      <td>540040</td>
      <td>IVORY LOVE BIRD CANDLE</td>
      <td>24</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.42</td>
      <td>12483.0</td>
      <td>Sweden</td>
      <td>4657</td>
      <td>10.08</td>
    </tr>
    <tr>
      <th>3</th>
      <td>540040</td>
      <td>BLACK LOVE BIRD CANDLE</td>
      <td>24</td>
      <td>2011-01-04 14:20:00</td>
      <td>0.42</td>
      <td>12483.0</td>
      <td>Sweden</td>
      <td>4657</td>
      <td>10.08</td>
    </tr>
    <tr>
      <th>4</th>
      <td>540040</td>
      <td>FELTCRAFT DOLL ROSIE</td>
      <td>12</td>
      <td>2011-01-04 14:20:00</td>
      <td>2.95</td>
      <td>12483.0</td>
      <td>Sweden</td>
      <td>4657</td>
      <td>35.40</td>
    </tr>
    <tr>
      <th>...</th>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
    </tr>
    <tr>
      <th>351001</th>
      <td>581585</td>
      <td>FAIRY TALE COTTAGE NIGHT LIGHT</td>
      <td>12</td>
      <td>2011-12-09 12:31:00</td>
      <td>1.95</td>
      <td>15804.0</td>
      <td>United Kingdom</td>
      <td>4318</td>
      <td>23.40</td>
    </tr>
    <tr>
      <th>351002</th>
      <td>581586</td>
      <td>LARGE CAKE STAND  HANGING STRAWBERY</td>
      <td>8</td>
      <td>2011-12-09 12:49:00</td>
      <td>2.95</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
      <td>4318</td>
      <td>23.60</td>
    </tr>
    <tr>
      <th>351003</th>
      <td>581586</td>
      <td>SET OF 3 HANGING OWLS OLLIE BEAK</td>
      <td>24</td>
      <td>2011-12-09 12:49:00</td>
      <td>1.25</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
      <td>4318</td>
      <td>30.00</td>
    </tr>
    <tr>
      <th>351004</th>
      <td>581586</td>
      <td>RED RETROSPOT ROUND CAKE TINS</td>
      <td>24</td>
      <td>2011-12-09 12:49:00</td>
      <td>8.95</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
      <td>4318</td>
      <td>214.80</td>
    </tr>
    <tr>
      <th>351005</th>
      <td>581586</td>
      <td>DOORMAT RED RETROSPOT</td>
      <td>10</td>
      <td>2011-12-09 12:49:00</td>
      <td>7.08</td>
      <td>13113.0</td>
      <td>United Kingdom</td>
      <td>4318</td>
      <td>70.80</td>
    </tr>
  </tbody>
</table>
<p>351006 rows × 9 columns</p>
</div>



### 고객별 RFM 구하기


```python
retail_customer_rfm = retail_orders_df.groupby("Customer ID").agg({"Recency":"min","Customer ID":"count","Monetary":"sum"})
retail_customer_rfm = retail_customer_rfm.rename(columns={"Customer ID":"Frequency"})
retail_customer_rfm = retail_customer_rfm.reset_index()
retail_customer_rfm
```




<div>

<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Customer ID</th>
      <th>Recency</th>
      <th>Frequency</th>
      <th>Monetary</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>12346.0</td>
      <td>4643</td>
      <td>1</td>
      <td>77183.60</td>
    </tr>
    <tr>
      <th>1</th>
      <td>12349.0</td>
      <td>4336</td>
      <td>72</td>
      <td>1457.55</td>
    </tr>
    <tr>
      <th>2</th>
      <td>12354.0</td>
      <td>4550</td>
      <td>58</td>
      <td>1079.40</td>
    </tr>
    <tr>
      <th>3</th>
      <td>12356.0</td>
      <td>4340</td>
      <td>58</td>
      <td>2487.43</td>
    </tr>
    <tr>
      <th>4</th>
      <td>12358.0</td>
      <td>4319</td>
      <td>17</td>
      <td>928.06</td>
    </tr>
    <tr>
      <th>...</th>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
    </tr>
    <tr>
      <th>4100</th>
      <td>18280.0</td>
      <td>4595</td>
      <td>10</td>
      <td>180.60</td>
    </tr>
    <tr>
      <th>4101</th>
      <td>18281.0</td>
      <td>4498</td>
      <td>7</td>
      <td>80.82</td>
    </tr>
    <tr>
      <th>4102</th>
      <td>18282.0</td>
      <td>4325</td>
      <td>12</td>
      <td>178.05</td>
    </tr>
    <tr>
      <th>4103</th>
      <td>18283.0</td>
      <td>4321</td>
      <td>719</td>
      <td>2039.58</td>
    </tr>
    <tr>
      <th>4104</th>
      <td>18287.0</td>
      <td>4360</td>
      <td>70</td>
      <td>1837.28</td>
    </tr>
  </tbody>
</table>
<p>4105 rows × 4 columns</p>
</div>



### 정규화(Normalization)
값의 범위를 0~1사이로 변환시켜 모든 컬럼의 데이터가 평등하게 만들어준다.  
서로 다른 단위의 값은 비교대상이 될 수 없다. 예를 들어, 80kg과 180cm는 비교할 수 없기에 정규화를 사용하여 비교한다.


```python
# 최대값과 최소값의 범위를 정해서 정규화할 수 있는 머신러닝
from sklearn.preprocessing import MinMaxScaler

#  공식 (X - MIN) / (MAX - MIN)

rfm = ["Recency","Frequency","Monetary"]

normalization = MinMaxScaler()
rfm_normalization = normalization.fit_transform(retail_customer_rfm[rfm])
rfm_normalization = pd.DataFrame(rfm_normalization, columns=rfm)

display(rfm_normalization)
```


<div>

<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Recency</th>
      <th>Frequency</th>
      <th>Monetary</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>0.958702</td>
      <td>0.000000</td>
      <td>0.285262</td>
    </tr>
    <tr>
      <th>1</th>
      <td>0.053097</td>
      <td>0.009625</td>
      <td>0.005373</td>
    </tr>
    <tr>
      <th>2</th>
      <td>0.684366</td>
      <td>0.007727</td>
      <td>0.003976</td>
    </tr>
    <tr>
      <th>3</th>
      <td>0.064897</td>
      <td>0.007727</td>
      <td>0.009180</td>
    </tr>
    <tr>
      <th>4</th>
      <td>0.002950</td>
      <td>0.002169</td>
      <td>0.003416</td>
    </tr>
    <tr>
      <th>...</th>
      <td>...</td>
      <td>...</td>
      <td>...</td>
    </tr>
    <tr>
      <th>4100</th>
      <td>0.817109</td>
      <td>0.001220</td>
      <td>0.000654</td>
    </tr>
    <tr>
      <th>4101</th>
      <td>0.530973</td>
      <td>0.000813</td>
      <td>0.000285</td>
    </tr>
    <tr>
      <th>4102</th>
      <td>0.020649</td>
      <td>0.001491</td>
      <td>0.000644</td>
    </tr>
    <tr>
      <th>4103</th>
      <td>0.008850</td>
      <td>0.097330</td>
      <td>0.007525</td>
    </tr>
    <tr>
      <th>4104</th>
      <td>0.123894</td>
      <td>0.009353</td>
      <td>0.006777</td>
    </tr>
  </tbody>
</table>
<p>4105 rows × 3 columns</p>
</div>


### 데이터 마이닝
- 대규모로 저장된 데이터안에서 체계적이고 자동적으로 통계적 규칙이나 짜임 또는 패턴을 분석하여, 가치있는 정보를 빼내는 과정이다.

### 클러스터 분석(Cluster analysis)
- 주어진 데이터들의 특성을 고려해 데이터 집단을 정의하고 데이터 집단을 대표할 수 있는 대표점을 찾는 것으로 데이터 마이닝의 한 방법이다.
- 클러스터란 비슷한 특성을 가진 데이터들의 집단이고 데이터의 성격에 따라 여러 클러스터(집단)으로 나뉠 수 있다.
- 머신러닝을 사용하여 분석할 수 있으며, 분리할 클러스터 개수를 알려줘야 한다.  
이 때, KElbowVisualizer를 사용해서 최적의 클러스터 개수(elbow)를 먼저 알아낼 수 있다.

### 병합 집단 알고리즘(Agglomerative clustering)
- 점과 점 사이의 거리값을 사용하여 샘플 데이터의 클러스터 쌍을 반복적으로 병합한다.
- 전달받은 클러스터 개수와 일치할 때 종료한다.

![image](https://github.com/SilverSharkDeveloper/Data-Analysis/assets/129861299/18b148ef-3f82-40dd-8fcd-4572a404f365)


```python
# conda install -c districtdatalabs yellowbrick
from yellowbrick.cluster import KElbowVisualizer
from sklearn.cluster import AgglomerativeClustering

model = AgglomerativeClustering()

# k_elbow_visualizer = KElbowVisualizer(model, k=(3, 9), timings=True)
k_elbow_visualizer = KElbowVisualizer(model, k=(2, 9), timings=False)
k_elbow_visualizer.fit(rfm_normalization)

k_elbow_visualizer.show()

#->score는 집단간의 응집도를 나타내고 k는 집단갯수를 나타내는데 응집도는 낮을수록 군집화가 잘되고, 집단갯수는 많을수록 좋지만 
#집단이 많아지면 모델 학습시 성능이 매우 떨어지므로 기울기가 확 줄기전 즉 집단의 갯수가 늘어난다고해서 응집도가 크게 줄어들지 않는 지점
#그 지점에서의 k값이 분류할 클러스터 갯수로 적합하다고 판단한다.
```

![output_25_0](https://github.com/SilverSharkDeveloper/Data-Analysis/assets/129861299/4c6b5321-c2ee-4264-bcb5-1996bfff7e35)







    <Axes: title={'center': 'Distortion Score Elbow for AgglomerativeClustering Clustering'}, xlabel='k', ylabel='distortion score'>




```python
#클러스터 4개로 군집화
agglomerative_clustering = AgglomerativeClustering(n_clusters=4).fit(rfm_normalization)

#결과 값 rfm df에 넣고 확인하기
rfm_normalization.loc[:, 'cluster'] = agglomerative_clustering.labels_
rfm_normalization.loc[:, 'cluster'].value_counts()
```




    cluster
    0    2510
    1     649
    2     489
    3     457
    Name: count, dtype: int64



### 시각화
- matplotlib를 사용하여 데이터를 시각화할 수 있고 이를 통해 한 눈에 볼 수 없는 많은 데이터를 한 눈에 볼 수 있다.
- 데이터 분석에 대한 전문 지식이 없는 일반인도 이해할 수 있다.
- 동일한 통계를 가지고 있더라도, 시각화 시 변화나 패턴이 다를 수 있다.


```python
import matplotlib.pyplot as plt
import seaborn as sns

# Recency : 얼마나 최근에 구매했는가
# Frequency : 얼마나 자주 구매했는가
# Monetary : 얼마나 많은 금액을 지출했는가
titles = ['Recency', 'Frequency', 'Monetary']

# 클러스터(집단) 개수
k = 4

# 각 항목별
for title in titles:
    plt.figure(figsize=(5, 5))
    
    # 집단 별
    for i in range(k):
#     scatter: 산점도(분포도)
        plt.scatter(rfm_normalization.loc[rfm_normalization['cluster'] == i, 'cluster'], 
                    rfm_normalization.loc[rfm_normalization['cluster'] == i, title], 
                    label=f'cluster {i}')

#     색상별 제목(label) 표시
    plt.legend()
    plt.title(title, size=15)
    plt.xlabel('cluster', size=12)
    plt.ylabel(title, size=12)
    plt.show()   
```


    

    
![output_28_0](https://github.com/SilverSharkDeveloper/Data-Analysis/assets/129861299/efa1b375-3db0-473f-b60c-3b86473bf23f)


![output_28_1](https://github.com/SilverSharkDeveloper/Data-Analysis/assets/129861299/01c9070d-54b9-460a-9d22-1d20c90d21cc)

    
![output_28_2](https://github.com/SilverSharkDeveloper/Data-Analysis/assets/129861299/73b93560-9d98-4ed0-bda0-06e26601d004)


### boxplot
- 데이터의 분포와 이상치를 한 번에 볼 수 있으며, 서로 다른 집단을 쉽게 비교할 수 있는 차트이다.
<div style="display: flex">
    <img src="./images/boxplot2.png" width="400" style="margin-left: 20px">
    <img src="./images/boxplot1.png" width="400" style="margin-left: 20px">
<div>
    
> 상자(Box): 상자의 아랫부분(Q1 또는 25th 백분위수)와 윗부분(Q3 또는 75th 백분위수)은 데이터의 중간 50% 범위를 나타냅니다. 상자의 높이는 이 범위를 나타냅니다.  

> 수염(Whiskers): 상자 위와 아래에 있는 선은 데이터의 최솟값과 최댓값을 나타냅니다. 보통, 수염의 길이는 박스의 높이의 1.5배 범위 이내의 값들로 정해집니다. 범위를 벗어나는 데이터는 이상치로 표시됩니다.  

> 중앙값(Median): 상자 내부에 있는 가로선은 데이터의 중앙값을 나타냅니다.  

> 이상치(Outliers): 수염 바깥에 있는 점들은 이상치로 간주되며, 주어진 데이터 집합에서 다른 값들과 크게 다른 데이터 포인트를 나타냅니다.


```python
titles = ['Recency', 'Frequency', 'Monetary']

for title in titles:
    plt.figure(figsize=(5, 5))
    sns.boxplot(x=rfm_normalization.cluster, y=rfm_normalization[title], palette='muted')
    plt.title(title)
    plt.show()
```

![output_30_0](https://github.com/SilverSharkDeveloper/Data-Analysis/assets/129861299/16ba934d-81fe-46c8-a2ab-33e324a7fa08)

![output_30_1](https://github.com/SilverSharkDeveloper/Data-Analysis/assets/129861299/09288419-1787-4cd8-88f3-eb3f5d956a4a)

![output_30_2](https://github.com/SilverSharkDeveloper/Data-Analysis/assets/129861299/7179d498-4541-4d00-b134-f3b9676c71a0)


### 박스 그래프에서 이상치로 판단한 값은 모두 중앙값으로 대체


```python
for i in range(4):
    rfm_normalization.loc[rfm_normalization[(
        rfm_normalization["cluster"]==i)&(
        rfm_normalization["Frequency"]>0.035)].index,"Frequency"]= \
    rfm_normalization[rfm_normalization["cluster"]==i]["Frequency"].median()
    rfm_normalization.loc[rfm_normalization[(
        rfm_normalization["cluster"]==i)&(
        rfm_normalization["Monetary"]>0.02)].index,"Monetary"]=\
    rfm_normalization[rfm_normalization["cluster"]==i]["Monetary"].median()

```


```python
titles = ['Recency', 'Frequency', 'Monetary']

for title in titles:
    plt.figure(figsize=(5, 5))
    sns.boxplot(x=rfm_normalization.cluster, y=rfm_normalization[title], palette='muted')
    plt.title(title)
    plt.show()
```

![output_33_0](https://github.com/SilverSharkDeveloper/Data-Analysis/assets/129861299/a050e225-a3a5-45d5-98c3-d5ceb090a6b5)


![output_33_1](https://github.com/SilverSharkDeveloper/Data-Analysis/assets/129861299/7dc3678e-6d62-4ebb-8845-56646428b3de)


![output_33_2](https://github.com/SilverSharkDeveloper/Data-Analysis/assets/129861299/510dd81a-06e9-4897-8a79-ddd84f09fa12)

- R, F, M 점수는 1~4로 계산함.
- 여기서는 최근 구매를 더 중요하다고 가정한다.
> Cluster0: VIP, (4, 4, 4) = 12  
Cluster1: GOLD, (3, 3, 3) = 9  
Cluster2: BRONZE, (1, 1, 1) = 3  
Cluster3: SILVER (2, 2, 2) = 6  

### 원본 고객 rfm 데이터에 정보 추가


```python
retail_customer_rfm['cluster'] = rfm_normalization['cluster'].replace([0, 1, 2, 3], ['VIP', 'GOLD', 'BRONZE', 'SILVER'])
retail_customer_rfm
```




<div>

<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Customer ID</th>
      <th>Recency</th>
      <th>Frequency</th>
      <th>Monetary</th>
      <th>cluster</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>12346.0</td>
      <td>4643</td>
      <td>1</td>
      <td>77183.60</td>
      <td>BRONZE</td>
    </tr>
    <tr>
      <th>1</th>
      <td>12349.0</td>
      <td>4336</td>
      <td>72</td>
      <td>1457.55</td>
      <td>VIP</td>
    </tr>
    <tr>
      <th>2</th>
      <td>12354.0</td>
      <td>4550</td>
      <td>58</td>
      <td>1079.40</td>
      <td>BRONZE</td>
    </tr>
    <tr>
      <th>3</th>
      <td>12356.0</td>
      <td>4340</td>
      <td>58</td>
      <td>2487.43</td>
      <td>VIP</td>
    </tr>
    <tr>
      <th>4</th>
      <td>12358.0</td>
      <td>4319</td>
      <td>17</td>
      <td>928.06</td>
      <td>VIP</td>
    </tr>
    <tr>
      <th>...</th>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
    </tr>
    <tr>
      <th>4100</th>
      <td>18280.0</td>
      <td>4595</td>
      <td>10</td>
      <td>180.60</td>
      <td>BRONZE</td>
    </tr>
    <tr>
      <th>4101</th>
      <td>18281.0</td>
      <td>4498</td>
      <td>7</td>
      <td>80.82</td>
      <td>SILVER</td>
    </tr>
    <tr>
      <th>4102</th>
      <td>18282.0</td>
      <td>4325</td>
      <td>12</td>
      <td>178.05</td>
      <td>VIP</td>
    </tr>
    <tr>
      <th>4103</th>
      <td>18283.0</td>
      <td>4321</td>
      <td>719</td>
      <td>2039.58</td>
      <td>VIP</td>
    </tr>
    <tr>
      <th>4104</th>
      <td>18287.0</td>
      <td>4360</td>
      <td>70</td>
      <td>1837.28</td>
      <td>VIP</td>
    </tr>
  </tbody>
</table>
<p>4105 rows × 5 columns</p>
</div>



### RFM 클러스터 데이터 시각화


```python
order = ['VIP', 'GOLD', 'BRONZE', 'SILVER']
sns.countplot(x='cluster', data=retail_customer_rfm, palette='muted', order=order)
```




    <Axes: xlabel='cluster', ylabel='count'>



![output_38_1](https://github.com/SilverSharkDeveloper/Data-Analysis/assets/129861299/f704029d-48c1-44f2-8d92-c1c81c253a07)


    


## 고객 정보 데이터가 있는 테이블이 있다면 조인해 고객의 나이, 결혼여부, 자녀 수 등으로 다양하게 클러스터를 시각화 할 수 있다.
