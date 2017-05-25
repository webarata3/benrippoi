package link.webarata3.poi;

import org.junit.Rule;
import org.junit.experimental.runners.Enclosed;
import org.junit.experimental.theories.DataPoints;
import org.junit.experimental.theories.Theories;
import org.junit.experimental.theories.Theory;
import org.junit.rules.ExpectedException;
import org.junit.rules.TemporaryFolder;
import org.junit.runner.RunWith;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;

import static org.hamcrest.Matchers.closeTo;
import static org.hamcrest.Matchers.is;
import static org.junit.Assert.assertThat;

@RunWith(Enclosed.class)
public class CellProxyTest {
    @RunWith(Theories.class)
    public static class 正常系_toStr {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("B2", "あいうえお"),
            new Fixture("C2", "123"),
            new Fixture("D2", "150.51"),
            new Fixture("F2", "true"),
            new Fixture("G2", "123150.51"),
            new Fixture("H2", ""),
            new Fixture("I2", ""),
            new Fixture("J2", "あいうえお123")
        };

        static class Fixture {
            String cellLabel;
            String expected;

            Fixture(String cellLabel, String expected) {
                this.cellLabel = cellLabel;
                this.expected = expected;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    ", expected='" + expected + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            CellProxy cellProxy = TestUtil.getCellProxy(tempFolder, "book1.xlsx", fixture.cellLabel);
            assertThat(cellProxy.toStr(), is(fixture.expected));
        }
    }

    @RunWith(Theories.class)
    public static class 異常系_日付仮_toStr {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();
        @Rule
        public ExpectedException thrown = ExpectedException.none();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("E2")
        };

        static class Fixture {
            String cellLabel;

            Fixture(String cellLabel) {
                this.cellLabel = cellLabel;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            CellProxy cellProxy = TestUtil.getCellProxy(tempFolder, "book1.xlsx", fixture.cellLabel);
            thrown.expect(UnsupportedOperationException.class);
            cellProxy.toStr();
        }
    }

    @RunWith(Theories.class)
    public static class 異常系_toStr {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();
        @Rule
        public ExpectedException thrown = ExpectedException.none();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("K2")
        };

        static class Fixture {
            String cellLabel;

            Fixture(String cellLabel) {
                this.cellLabel = cellLabel;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            CellProxy cellProxy = TestUtil.getCellProxy(tempFolder, "book1.xlsx", fixture.cellLabel);
            thrown.expect(PoiIllegalAccessException.class);
            cellProxy.toStr();
        }
    }

    @RunWith(Theories.class)
    public static class 正常系_toInt {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("B3", 456),
            new Fixture("C3", 123),
            new Fixture("D3", 150),
            new Fixture("G3", 369),
            new Fixture("J3", 456123)
        };

        static class Fixture {
            String cellLabel;
            int expected;

            Fixture(String cellLabel, int expected) {
                this.cellLabel = cellLabel;
                this.expected = expected;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    ", expected='" + expected + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            CellProxy cellProxy = TestUtil.getCellProxy(tempFolder, "book1.xlsx", fixture.cellLabel);
            assertThat(cellProxy.toInt(), is(fixture.expected));
        }
    }

    @RunWith(Theories.class)
    public static class 異常系_toInt {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();
        @Rule
        public ExpectedException thrown = ExpectedException.none();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("B2"),
            new Fixture("E3"),
            new Fixture("F3"),
            new Fixture("H3"),
            new Fixture("I3"),
            new Fixture("K3")
        };

        static class Fixture {
            String cellLabel;

            Fixture(String cellLabel) {
                this.cellLabel = cellLabel;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            CellProxy cellProxy = TestUtil.getCellProxy(tempFolder, "book1.xlsx", fixture.cellLabel);
            thrown.expect(PoiIllegalAccessException.class);
            cellProxy.toInt();
        }
    }

    @RunWith(Theories.class)
    public static class 正常系_toDouble {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("B4", 123.456),
            new Fixture("C4", 123),
            new Fixture("D4", 150.51),
            new Fixture("G4", 50.17),
            new Fixture("J4", 123123.456)
        };

        static class Fixture {
            String cellLabel;
            double expected;

            Fixture(String cellLabel, double expected) {
                this.cellLabel = cellLabel;
                this.expected = expected;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    ", expected='" + expected + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            CellProxy cellProxy = TestUtil.getCellProxy(tempFolder, "book1.xlsx", fixture.cellLabel);
            assertThat(cellProxy.toDouble(), is(closeTo(fixture.expected, 0.00001)));
        }
    }

    @RunWith(Theories.class)
    public static class 異常系_toDouble {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();
        @Rule
        public ExpectedException thrown = ExpectedException.none();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("B2"),
            new Fixture("E4"),
            new Fixture("F4"),
            new Fixture("H4"),
            new Fixture("I4"),
            new Fixture("K4")
        };

        static class Fixture {
            String cellLabel;

            Fixture(String cellLabel) {
                this.cellLabel = cellLabel;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            CellProxy cellProxy = TestUtil.getCellProxy(tempFolder, "book1.xlsx", fixture.cellLabel);
            thrown.expect(PoiIllegalAccessException.class);
            cellProxy.toDouble();
        }
    }

    @RunWith(Theories.class)
    public static class 正常系_toBoolean {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("F5", true),
            new Fixture("G5", false)
        };

        static class Fixture {
            String cellLabel;
            boolean expected;

            Fixture(String cellLabel, boolean expected) {
                this.cellLabel = cellLabel;
                this.expected = expected;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    ", expected='" + expected + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            CellProxy cellProxy = TestUtil.getCellProxy(tempFolder, "book1.xlsx", fixture.cellLabel);
            assertThat(cellProxy.toBoolean(), is(fixture.expected));
        }
    }

    @RunWith(Theories.class)
    public static class 異常系_toBoolean {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();
        @Rule
        public ExpectedException thrown = ExpectedException.none();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("B5"),
            new Fixture("C5"),
            new Fixture("D5"),
            new Fixture("E5"),
            new Fixture("K5")
        };

        static class Fixture {
            String cellLabel;

            Fixture(String cellLabel) {
                this.cellLabel = cellLabel;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            CellProxy cellProxy = TestUtil.getCellProxy(tempFolder, "book1.xlsx", fixture.cellLabel);
            thrown.expect(PoiIllegalAccessException.class);
            cellProxy.toBoolean();
        }
    }

    @RunWith(Theories.class)
    public static class 正常系_toLocalDate_日付 {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("E6", LocalDate.of(2015, 12, 1)),
            new Fixture("G6", LocalDate.of(2015, 12, 3))
        };

        static class Fixture {
            String cellLabel;
            LocalDate expected;

            Fixture(String cellLabel, LocalDate expected) {
                this.cellLabel = cellLabel;
                this.expected = expected;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    ", expected='" + expected + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            CellProxy cellProxy = TestUtil.getCellProxy(tempFolder, "book1.xlsx", fixture.cellLabel);
            assertThat(cellProxy.toLocalDate(), is(fixture.expected));
        }
    }

    @RunWith(Theories.class)
    public static class 異常_toLocalDate_日付 {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();
        @Rule
        public ExpectedException thrown = ExpectedException.none();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("A6"),
            new Fixture("B6"),
            new Fixture("C6"),
            new Fixture("D6"),
            new Fixture("F6"),
            new Fixture("H6"),
            new Fixture("I6"),
            new Fixture("J6"),
            new Fixture("K6")
        };

        static class Fixture {
            String cellLabel;

            Fixture(String cellLabel) {
                this.cellLabel = cellLabel;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            CellProxy cellProxy = TestUtil.getCellProxy(tempFolder, "book1.xlsx", fixture.cellLabel);
            thrown.expect(PoiIllegalAccessException.class);
            cellProxy.toLocalDate();
        }
    }

    @RunWith(Theories.class)
    public static class 正常系_toLocalTime_日付 {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("E7", LocalTime.of(10, 10, 30)),
            new Fixture("G7", LocalTime.of(12, 34, 30))
        };

        static class Fixture {
            String cellLabel;
            LocalTime expected;

            Fixture(String cellLabel, LocalTime expected) {
                this.cellLabel = cellLabel;
                this.expected = expected;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    ", expected='" + expected + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            CellProxy cellProxy = TestUtil.getCellProxy(tempFolder, "book1.xlsx", fixture.cellLabel);
            assertThat(cellProxy.toLocalTime(), is(fixture.expected));
        }
    }

    @RunWith(Theories.class)
    public static class 異常_toLocalTime_日付 {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();
        @Rule
        public ExpectedException thrown = ExpectedException.none();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("A7"),
            new Fixture("B7"),
            new Fixture("C7"),
            new Fixture("D7"),
            new Fixture("F7"),
            new Fixture("H7"),
            new Fixture("I7"),
            new Fixture("J7"),
            new Fixture("K7")
        };

        static class Fixture {
            String cellLabel;

            Fixture(String cellLabel) {
                this.cellLabel = cellLabel;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            CellProxy cellProxy = TestUtil.getCellProxy(tempFolder, "book1.xlsx", fixture.cellLabel);
            thrown.expect(PoiIllegalAccessException.class);
            cellProxy.toLocalTime();
        }
    }

    @RunWith(Theories.class)
    public static class 正常系_toLocalDateTime_日付 {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("E8", LocalDateTime.of(2015, 12, 1, 10, 10, 30)),
            new Fixture("G8", LocalDateTime.of(2015,12,3, 10,10, 30))
        };

        static class Fixture {
            String cellLabel;
            LocalDateTime expected;

            Fixture(String cellLabel, LocalDateTime expected) {
                this.cellLabel = cellLabel;
                this.expected = expected;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    ", expected='" + expected + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            CellProxy cellProxy = TestUtil.getCellProxy(tempFolder, "book1.xlsx", fixture.cellLabel);
            assertThat(cellProxy.toLocalDateTime(), is(fixture.expected));
        }
    }

    @RunWith(Theories.class)
    public static class 異常_toLocalDateTime_日付 {
        @Rule
        public TemporaryFolder tempFolder = new TemporaryFolder();
        @Rule
        public ExpectedException thrown = ExpectedException.none();

        @DataPoints
        public static Fixture[] PARAMs = {
            new Fixture("A8"),
            new Fixture("B8"),
            new Fixture("C8"),
            new Fixture("D8"),
            new Fixture("F8"),
            new Fixture("H8"),
            new Fixture("I8"),
            new Fixture("J8"),
            new Fixture("K8")
        };

        static class Fixture {
            String cellLabel;

            Fixture(String cellLabel) {
                this.cellLabel = cellLabel;
            }

            @Override
            public String toString() {
                return "Fixture{" +
                    "cellLabel='" + cellLabel + '\'' +
                    '}';
            }
        }

        @Theory
        public void test(Fixture fixture) throws Exception {
            CellProxy cellProxy = TestUtil.getCellProxy(tempFolder, "book1.xlsx", fixture.cellLabel);
            thrown.expect(PoiIllegalAccessException.class);
            cellProxy.toLocalDateTime();
        }
    }
}
