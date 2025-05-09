# PowerPoint 템플릿 문법 작성 지침

## 기본 원칙

1. **DollarSignEngine 친화적**: 기존 DollarSignEngine 사용 경험과 최대한 유사하게 유지
2. **역할 분리**: 슬라이드 노트는 제어 지시문용, 텍스트 요소는 값 바인딩용
3. **간결함**: 직관적이고 간소화된 문법으로 사용 편의성 극대화
4. **디자인 보존**: PPT 디자인 요소의 원래 의도 존중

## 문법 구조

### 1. 값 바인딩 (슬라이드 요소 내)
```
${속성명}                   // 기본 속성 바인딩
${객체.속성명}              // 중첩 속성 바인딩
${값:포맷지정자}            // 포맷 지정자 사용
${조건 ? 값1 : 값2}         // 조건부 표현식
${메소드()}                 // 메소드 호출
```

### 2. 특수 함수 (슬라이드 요소 내)
```
${ppt.Image("이미지속성")}   // 이미지 바인딩
${ppt.Chart("차트데이터")}   // 차트 데이터 바인딩
${ppt.Table("테이블데이터")} // 표 데이터 바인딩
```

### 3. 제어 지시문 (슬라이드 노트에만 배치)
```
#foreach: 컬렉션명 as 항목명, 옵션...    // 슬라이드 복제 (항목명 명시적 지정)
#foreach: 컬렉션명, 옵션...             // 슬라이드 복제 (항목명 자동 결정)
#if: 조건식, 옵션...                   // 조건부 처리
```

## 상세 문법 설명

### 1. 값 바인딩 (슬라이드 요소 내)

모든 텍스트 요소에서 DollarSignEngine의 문법을 그대로 사용:

```
제목: ${report.title}
날짜: ${DateTime.Now:yyyy-MM-dd}
합계: ${items.Sum(i => i.price):C2}
상태: ${status == "active" ? "활성" : "비활성"}
```

### 2. 특수 함수 (슬라이드 요소 내)

#### 이미지 바인딩
이미지 도형의 텍스트에:
```
${ppt.Image("company.logo")}
${ppt.Image("product.photo", width: 300, height: 200, preserveAspectRatio: true)}
```

#### 차트 데이터 바인딩
차트 도형의 텍스트에:
```
${ppt.Chart("salesData")}
${ppt.Chart("salesData", series: "series", categories: "categories", title: "월별 판매량")}
```

#### 표 데이터 바인딩
표 도형의 텍스트에:
```
${ppt.Table("employeeData")}
${ppt.Table("employeeData", headers: true, startRow: 1, endRow: 10)}
```

### 3. 제어 지시문 (슬라이드 노트에만 배치)

#### 슬라이드 복제 지시문 (#foreach)
```
#foreach: categories
#foreach: products as product
```

- 슬라이드의 디자인을 유지하면서 데이터 컬렉션의 각 항목에 대해 슬라이드를 복제합니다.
- 슬라이드 내부에 인덱스 참조(예: `${item[0]}`, `${item[1]}` 등)가 있는 경우, 각 슬라이드는 해당 인덱스 범위의 데이터로 채워집니다.
- 데이터 컬렉션의 크기에 따라 필요한 만큼 슬라이드가 생성됩니다.

예: 슬라이드에 `${item[0]}` ~ `${item[4]}`까지 참조가 있고 데이터가 12개인 경우:
- 첫 번째 슬라이드: `item[0]`~`item[4]`는 데이터의 0~4번 항목으로 채워짐
- 두 번째 슬라이드: `item[0]`~`item[4]`는 데이터의 5~9번 항목으로 채워짐 
- 세 번째 슬라이드: `item[0]`~`item[1]`은 데이터의 10~11번 항목으로 채워지고, `item[2]`~`item[4]`는 빈 값으로 처리됨

#### 조건부 지시문
```
#if: report.hasChart, target: "sales_chart"
#if: total > 1000, target: "warning_box", visibleWhenFalse: "success_box"
```
- `target`: 조건부로 표시/숨김 처리할 도형의 이름
- `visibleWhenFalse`: 조건이 거짓일 때 표시할 대체 도형의 이름

## 변수명 자동결정 규칙

컬렉션 변수명에서 항목 변수명을 자동으로 결정하는 규칙:

1. **기본 규칙**: 컬렉션 이름에서 마지막 's'를 제거한 형태가 항목 변수명이 됩니다.
   - `Items` → `item`
   - `Products` → `product`
   - `Categories` → `category`

2. **예외 처리**: 's'로 끝나지 않는 컬렉션 이름이나 불규칙한 복수형의 경우, 컬렉션 이름을 소문자화하여 사용합니다.
   - `People` → `people`
   - `Data` → `data`

3. **명시적 지정**: `as` 키워드를 사용하여 변수명을 명시적으로 지정할 수 있습니다.
   - `#foreach: Products as p`에서 `${p.name}`으로 접근

## 여러 컬렉션 동시 처리

여러 컬렉션을 동시에 처리하는 방법:

```
#foreach: Products as product, Categories as category
```

이 경우, 각 슬라이드에서 `${product.name}`과 `${category.name}`으로 현재 항목에 접근할 수 있습니다.

## 문법 적용 위치

1. **값 바인딩 & 특수 함수**: 
   - 텍스트 상자, 도형, 표, 차트 등 슬라이드 요소의 텍스트 내용에 배치
   - PowerPoint에서 해당 요소를 선택하고 텍스트 편집 모드에서 입력

2. **제어 지시문**: 
   - 슬라이드 노트에만 배치 (보기 > 노트)
   - 여러 지시문이 필요한 경우 각각 새 줄에 배치

3. **요소 식별**: 
   - PowerPoint에서 도형 선택 → 오른쪽 클릭 → 이름 지정
   - 지정된 이름으로 제어 지시문에서 참조

## 슬라이드 노트 작성 예시

```
# 이 슬라이드는 제품 목록을 표시합니다
#foreach: products as item
#if: products.Count > 0, target: "products_container", visibleWhenFalse: "no_products_message"
```

## 예제 시나리오

### 기본 프레젠테이션 슬라이드

**슬라이드 요소 내용:**
- 제목 텍스트 상자: `${report.title}`
- 부제목 텍스트 상자: `${report.subtitle}`
- 날짜 텍스트 상자: `${report.date:yyyy-MM-dd}`
- 로고 이미지: `${ppt.Image("company.logo")}`

**슬라이드 노트:**
```
#if: report.isConfidential, target: "confidential_watermark"
```

### 제품 목록 슬라이드 (배열 인덱스 사용)

**슬라이드 요소 내용:**
- 제목 텍스트 상자: `${category.name} 제품 목록`
- 제품 항목 1: 
  ```
  ${item[0].Id}. ${item[0].Name} - ${item[0].Description}
  가격: ${item[0].Price:C0}원
  ```
- 제품 항목 2: 
  ```
  ${item[1].Id}. ${item[1].Name} - ${item[1].Description}
  가격: ${item[1].Price:C0}원
  ```
- 제품 항목 3-5: (유사한 형식으로 계속)

**슬라이드 노트:**
```
#foreach: categories as category
  #if: category.products.Length > 0, target: "products_container"
```

### 데이터 대시보드 슬라이드

**슬라이드 요소 내용:**
- 제목 텍스트 상자: `${period} 판매 분석`
- 차트 도형 (이름: "sales_chart"): 
  ```
  ${ppt.Chart("salesData", title: "${period} 판매 추이")}
  ```
- 표 도형 (이름: "top_products"): 
  ```
  ${ppt.Table("topProducts", headers: true)}
  ```

**슬라이드 노트:**
```
#if: salesData.Length > 0, target: "sales_chart", visibleWhenFalse: "no_data_message"
```

### 부서별 보고서 슬라이드 (여러 슬라이드 생성)

**슬라이드 요소 내용:**
- 제목 텍스트 상자: `${dept.name} 부서`
- 부서장 텍스트 상자: `부서장: ${dept.manager}`
- 인원수 텍스트 상자: `인원: ${dept.members.Length}명`
- 실적 차트 (이름: "dept_chart"): 
  ```
  ${ppt.Chart("dept.performance", title: "${dept.name} 부서 실적")}
  ```
- 팀원 1: `${member[0].name} (${member[0].position})`
- 팀원 2: `${member[1].name} (${member[1].position})`
- 팀원 3-8: (유사한 형식으로 계속)

**슬라이드 노트:**
```
#foreach: departments as dept
```

### 여러 컬렉션 동시 표시 (관련 데이터)

**슬라이드 요소 내용:**
- 제목 텍스트 상자: `${category.name} - ${product.name}`
- 제품 설명: `${product.description}`
- 카테고리 설명: `${category.description}`

**슬라이드 노트:**
```
#foreach: Products as product, Categories as category
```

이 예제에서는 Products와 Categories 배열을 동시에 처리하며, 인덱스가 같은 항목끼리 매칭됩니다.