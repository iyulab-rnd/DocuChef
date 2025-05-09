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
#foreach: 컬렉션명 as 항목명, 옵션...            // 반복 처리 (항목명 명시적 지정)
#foreach: 컬렉션명, 옵션...                     // 반복 처리 (항목명 자동 결정)
#if: 조건식, 옵션...                           // 조건부 처리
#slide-foreach: 컬렉션명 as 항목명, 옵션...     // 슬라이드 복제 (항목명 명시적 지정)
#slide-foreach: 컬렉션명, 옵션...              // 슬라이드 복제 (항목명 자동 결정)
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

#### 반복 지시문
```
#foreach: products, target: "product_box", maxItems: 6
#foreach: testimonials as testimonial, target: "quote_text"
```

- `target`: 반복적으로 복제할 도형/텍스트 상자의 이름
- `maxItems`: 한 슬라이드에 표시할 최대 항목 수
- `layout`: 레이아웃 유형 (grid, vertical, horizontal, etc.)
- `continueOnNewSlide`: 최대 항목 초과 시 새 슬라이드에 계속 (true/false)

#### 조건부 지시문
```
#if: report.hasChart, target: "sales_chart"
#if: total > 1000, target: "warning_box", visibleWhenFalse: "success_box"
```
- `target`: 조건부로 표시/숨김 처리할 도형의 이름
- `visibleWhenFalse`: 조건이 거짓일 때 표시할 대체 도형의 이름

#### 슬라이드 복제 지시문
```
#slide-foreach: categories
#slide-foreach: products as product, titleTarget: "product_title", imageTarget: "product_image"
#slide-foreach: departments as dept, maxItems: 10
```
- `titleTarget`: 제목을 표시할 텍스트 상자의 이름
- `imageTarget`: 이미지를 표시할 이미지 도형의 이름
- `maxItems`: 슬라이드당 표시할 최대 항목 수 (이 수를 초과하면 새 슬라이드 생성)
- `maxSlides`: 생성할 최대 슬라이드 수

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
   - `#slide-foreach: Products as p`에서 `${p.name}`으로 접근

## 여러 컬렉션 동시 처리

여러 컬렉션을 동시에 처리하는 방법:

```
#slide-foreach: Products as product, Categories as category, maxItems: 5
```

이 경우, 각 슬라이드에서 `${product.name}`과 `${category.name}`으로 현재 항목에 접근할 수 있습니다.

## 중첩 컬렉션 처리

중첩된 컬렉션을 처리하는 방법:

```
#slide-foreach: Departments, maxItems: 1
  #foreach: department.Employees as emp, target: "EmployeeTemplate", maxItems: 10
```

이 경우, 각 부서당 하나의 슬라이드가 생성되고, 각 슬라이드 내에서 해당 부서의 직원 목록이 처리됩니다.

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
#foreach: products, target: "product_item", maxItems: 4, layout: "grid", continueOnNewSlide: true
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

### 제품 목록 슬라이드

**슬라이드 요소 내용:**
- 제목 텍스트 상자: `${category.name} 제품 목록`
- 제품 템플릿 텍스트 상자 (이름: "product_item"): 
  ```
  ${item.name}
  가격: ${item.price:C2}
  ${item.isNew ? "[신제품]" : ""}
  ```
- 제품 이미지 도형 (이름: "product_image"): 
  ```
  ${ppt.Image("item.imageUrl")}
  ```

**슬라이드 노트:**
```
#slide-foreach: categories as category, maxItems: 1
  #foreach: category.products as item, target: "product_item", maxItems: 6, layout: "grid(3,2)"
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
#if: salesData.length > 0, target: "sales_chart", visibleWhenFalse: "no_data_message"
```

### 부서별 보고서 슬라이드 (여러 슬라이드 생성)

**슬라이드 요소 내용:**
- 제목 텍스트 상자 (이름: "dept_title"): `${dept.name} 부서`
- 부서장 텍스트 상자: `부서장: ${dept.manager}`
- 인원수 텍스트 상자: `인원: ${dept.members.length}명`
- 실적 차트 (이름: "dept_chart"): 
  ```
  ${ppt.Chart("dept.performance", title: "${dept.name} 부서 실적")}
  ```

**슬라이드 노트:**
```
#slide-foreach: departments as dept, maxItems: 1
  #foreach: dept.members as member, target: "member_template", maxItems: 8
```

### 여러 컬렉션 동시 표시 (관련 데이터)

**슬라이드 요소 내용:**
- 제목 텍스트 상자: `${category.name} - ${product.name}`
- 제품 설명: `${product.description}`
- 카테고리 설명: `${category.description}`

**슬라이드 노트:**
```
#slide-foreach: Products as product, Categories as category, maxItems: 1
```

이 예제에서는 Products와 Categories 배열을 동시에 처리하며, 인덱스가 같은 항목끼리 매칭됩니다.