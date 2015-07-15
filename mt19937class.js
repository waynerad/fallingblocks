/*jslint bitwise: true */
/* Modified from http://www.math.sci.hiroshima-u.ac.jp/~m-mat/MT/VERSIONS/JAVASCRIPT/java-script.html */
function MersenneTwister19937() {
    "use strict";
    var N, M, MATRIXA, UPPERMASK, LOWERMASK, mt, mti, i;
    N = 624;
    M = 397;
    MATRIXA = 0x9908b0df;
    UPPERMASK = 0x80000000;
    LOWERMASK = 0x7fffffff;
    mt = [];
    for (i = 0; i < N; i = i + 1) {
        mt[i] = 0;
    }
    mti = N + 1;
    function unsigned32(n1) {
        return n1 < 0 ? (n1 ^ UPPERMASK) + UPPERMASK : n1;
    }
    function subtraction32(n1, n2) {
        return n1 < n2 ? unsigned32((0x100000000 - (n2 - n1)) & 0xffffffff) : n1 - n2;
    }
    function addition32(n1, n2) {
        return unsigned32((n1 + n2) & 0xffffffff);
    }
    function multiplication32(n1, n2) {
        var sum, i2;
        sum = 0;
        for (i2 = 0; i2 < 32; i2 = i2 + 1) {
            if ((n1 >>> i2) & 0x1) {
                sum = addition32(sum, unsigned32(n2 << i2));
            }
        }
        return sum;
    }
    this.initGenrand = function (s) {
        mt[0] = unsigned32(s & 0xffffffff);
        for (mti = 1; mti < N; mti = mti + 1) {
            mt[mti] = addition32(multiplication32(1812433253, unsigned32(mt[mti - 1] ^ (mt[mti - 1] >>> 30))), mti);
            mt[mti] = unsigned32(mt[mti] & 0xffffffff);
        }
    };
    this.initByArray = function (initKey, keyLength) {
        var i2, j, k, x;
        this.initGenrand(19650218);
        i2 = 1;
        j = 0;
        if (N > keyLength) {
            x = N;
        } else {
            x = keyLength;
        }
        for (k = x; k > 0; k = k - 1) {
            mt[i2] = addition32(addition32(unsigned32(mt[i2] ^ multiplication32(unsigned32(mt[i2 - 1] ^ (mt[i2 - 1] >>> 30)), 1664525)), initKey[j]), j);
            mt[i2] = unsigned32(mt[i2] & 0xffffffff);
            i2 = i2 + 1;
            j = j + 1;
            if (i2 >= N) {
                mt[0] = mt[N - 1];
                i2 = 1;
            }
            if (j >= keyLength) {
                j = 0;
            }
        }
        for (k = N - 1; k > 0; k = k - 1) {
            mt[i2] = subtraction32(unsigned32((mt[i2]) ^ multiplication32(unsigned32(mt[i2 - 1] ^ (mt[i2 - 1] >>> 30)), 1566083941)), i2);
            mt[i2] = unsigned32(mt[i2] & 0xffffffff);
            i2 = i2 + 1;
            if (i2 >= N) {
                mt[0] = mt[N - 1];
                i2 = 1;
            }
        }
        mt[0] = 0x80000000;
    };
    this.genrandInt32 = function () {
        var y, mag01, kk, x;
        mag01 = [0x0, MATRIXA];
        if (mti >= N) {
            if (mti === (N + 1)) {
                this.initGenrand(5489);
            }
            for (kk = 0; kk < (N - M); kk = kk + 1) {
                y = unsigned32((mt[kk] & UPPERMASK) | (mt[kk + 1] & LOWERMASK));
                mt[kk] = unsigned32(mt[kk + M] ^ (y >>> 1) ^ mag01[y & 0x1]);
            }
            x = kk;
            for (kk = x; kk < (N - 1); kk = kk + 1) {
                y = unsigned32((mt[kk] & UPPERMASK) | (mt[kk + 1] & LOWERMASK));
                mt[kk] = unsigned32(mt[kk + (M - N)] ^ (y >>> 1) ^ mag01[y & 0x1]);
            }
            y = unsigned32((mt[N - 1] & UPPERMASK) | (mt[0] & LOWERMASK));
            mt[N - 1] = unsigned32(mt[M - 1] ^ (y >>> 1) ^ mag01[y & 0x1]);
            mti = 0;
        }
        y = mt[mti];
        mti = mti + 1;
        y = unsigned32(y ^ (y >>> 11));
        y = unsigned32(y ^ ((y << 7) & 0x9d2c5680));
        y = unsigned32(y ^ ((y << 15) & 0xefc60000));
        y = unsigned32(y ^ (y >>> 18));
        return y;
    };
    /* generates a random number on [0,0x7fffffff]-interval */
    this.genrandInt31 = function () {
        return (this.genrandInt32() >>> 1);
    };
    /* generates a random number on [0,1]-real-interval */
    this.genrandReal1 = function () {
        return this.genrandInt32() * (1.0 / 4294967295.0);
    };
    /* generates a random number on [0,1)-real-interval */
    this.genrandReal2 = function () {
        return this.genrandInt32() * (1.0 / 4294967296.0);
    };
    /* generates a random number on (0,1)-real-interval */
    this.genrandReal3 = function () {
        return ((this.genrandInt32()) + 0.5) * (1.0 / 4294967296.0);
    };
    /* generates a random number on [0,1) with 53-bit resolution */
    this.genrandRes53 = function () {
        var a = this.genrandInt32() >>> 5, b = this.genrandInt32() >>> 6;
        return (a * 67108864.0 + b) * (1.0 / 9007199254740992.0);
    };
}
